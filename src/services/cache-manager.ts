export interface CacheData {
  clientes: Array<{ id: string; title: string }>;
  asuntos: Array<{ id: string; title: string; clienteId: string }>;
  subasuntos: Array<{ id: string; title: string; asuntoId: string }>;
  tiposDocumento: Array<{ id: string; title: string }>;
  subtipos: Array<{ id: string; title: string; tipoId: string }>;
  lastUpdated: string;
  version: string;
}

export interface CacheMetadata {
  lastUpdated: Date;
  version: string;
  recordCount: number;
  isStale: boolean;
}

export class CacheManager {
  private readonly CACHE_KEY = 'sharepoint_metadata_cache';
  private readonly CACHE_VERSION = '1.0.4';
  //   private readonly CACHE_EXPIRY_HOURS = 24;
  private readonly STALE_THRESHOLD_HOURS = 6;

  private cache: CacheData | null = null;
  private isUpdating = false;
  private originalDataService: any;
  private loadingPromises = new Map<string, Promise<Array<{ id: string; title: string }>>>();

  constructor(originalDataService: any) {
    this.originalDataService = originalDataService;
  }

  async initialize(): Promise<CacheData> {
    try {
      const cachedData = this.loadFromStorage();
      if (cachedData && this.isValidCache(cachedData)) {
        this.cache = cachedData;
        setTimeout(() => this.expandCacheInBackground(), 3000);
        return cachedData;
      } else {
        return await this.createInitialCache();
      }
    } catch (error) {
      return this.createEmptyCache();
    }
  }

  private createEmptyCache(): CacheData {
    return {
      clientes: [],
      asuntos: [],
      subasuntos: [],
      tiposDocumento: [],
      subtipos: [],
      lastUpdated: new Date().toISOString(),
      version: this.CACHE_VERSION
    };
  }

  private async createInitialCache(): Promise<CacheData> {
    if (this.isUpdating) {
      return this.createEmptyCache();
    }
    this.isUpdating = true;
    try {
      const [clientes, tiposDocumento] = await Promise.all([
        this.originalDataService.getClientesOriginal(),
        this.originalDataService.getTiposDocumento()
      ]);

      const initialCache: CacheData = {
        clientes: clientes.sort((a, b) => a.title.localeCompare(b.title)),
        asuntos: [],
        subasuntos: [],
        tiposDocumento: tiposDocumento.sort((a, b) => a.title.localeCompare(b.title)),
        subtipos: [],
        lastUpdated: new Date().toISOString(),
        version: this.CACHE_VERSION
      };

      this.saveToStorage(initialCache);
      this.cache = initialCache;
      return initialCache;
    } catch (error) {
      return this.createEmptyCache();
    } finally {
      this.isUpdating = false;
    }
  }

  getClientes(searchText: string = '', maxResults: number = 100): Array<{ id: string; title: string }> {
    if (!this.cache) return [];

    let results = this.cache.clientes;

    if (searchText) {
      const normalizedSearch = this.normalizeText(searchText);
      results = results.filter(cliente =>
        this.normalizeText(cliente.title).includes(normalizedSearch) ||
        this.normalizeText(cliente.id).includes(normalizedSearch)
      );
    }

    return results.slice(0, maxResults);
  }

  async getAsuntosByCliente(clienteId: string): Promise<Array<{ id: string; title: string }>> {
    if (!this.cache) return [];
    const cached = this.cache.asuntos.filter(asunto => asunto.clienteId === clienteId);
    if (cached.length > 0) {
      return cached.sort((a, b) => a.title.localeCompare(b.title));
    }
    if (this.loadingPromises.has(clienteId)) {
      return await this.loadingPromises.get(clienteId)!;
    }
    const loadingPromise = this.loadAsuntosForCliente(clienteId);
    this.loadingPromises.set(clienteId, loadingPromise);
    try {
      const asuntos = await loadingPromise;
      return asuntos;
    } finally {
      this.loadingPromises.delete(clienteId);
    }
  }

  private async loadAsuntosForCliente(clienteId: string): Promise<Array<{ id: string; title: string }>> {
    try {
      const cliente = this.cache?.clientes.find(c => c.id === clienteId);
      if (!cliente) {
        console.warn(`Cliente ${clienteId} no encontrado en cache`);
        return [];
      }

      const clienteAsuntos = await this.originalDataService.getAsuntosByClienteOriginal(clienteId, cliente.title);

      if (clienteAsuntos.length > 0) {
        const filteredAsuntos = clienteAsuntos.filter((asunto: any) => {
          // Include asuntos that match the specific cliente
          const clienteIdMatch = asunto.clienteId?.toString().trim() === clienteId.trim();
          const clienteNameMatch = asunto.clienteName?.toString().trim().toLowerCase() === cliente.title.toLowerCase().trim();
          const matchesSpecificCliente = clienteIdMatch && clienteNameMatch;
          
          // Include asuntos that have null cliente (should appear for all clientes)
          const hasNullCliente = (!asunto.clienteId || asunto.clienteId.toString().trim() === '') && 
                                (!asunto.clienteName || asunto.clienteName.toString().trim() === '');
          
          return matchesSpecificCliente || hasNullCliente;
        });
        const mappedAsuntos = filteredAsuntos.map((asunto: any) => ({
          id: asunto.id,
          title: asunto.title,
          clienteId: clienteId
        }));

        if (this.cache) {
          const updatedAsuntos = [...this.cache.asuntos, ...mappedAsuntos];
          const updatedCache: CacheData = {
            ...this.cache,
            asuntos: updatedAsuntos.sort((a, b) => a.title.localeCompare(b.title)),
            lastUpdated: new Date().toISOString()
          };
          this.saveToStorage(updatedCache);
          this.cache = updatedCache;
        }
        return mappedAsuntos.sort((a, b) => a.title.localeCompare(b.title));
      }
      return [];
    } catch (error) {
      console.error(`Error cargando asuntos para cliente ${clienteId}:`, error);
      return [];
    }
  }

  private async expandCacheInBackground(): Promise<void> {
    if (this.isUpdating || !this.cache) return;
    this.isUpdating = true;
    try {
      const clientesConAsuntos = new Set(this.cache.asuntos.map(a => a.clienteId));
      const clientesSinAsuntos = this.cache.clientes
        .filter(cliente => !clientesConAsuntos.has(cliente.id))
        .slice(0, 10);

      for (const cliente of clientesSinAsuntos) {
        try {
          await this.loadAsuntosForCliente(cliente.id);
          await new Promise(resolve => setTimeout(resolve, 100));
        } catch (error) { }
      }
    } catch (error) {
      console.error('Error expandiendo cache:', error);
    } finally {
      this.isUpdating = false;
    }
  }

  getSubasuntosByAsunto(asuntoId: string): Array<{ id: string; title: string }> {
    if (!this.cache) return [];

    return this.cache.subasuntos
      .filter(subasunto => subasunto.asuntoId === asuntoId)
      .sort((a, b) => a.title.localeCompare(b.title));
  }

  getTiposDocumento(): Array<{ id: string; title: string }> {
    if (!this.cache) return [];
    return this.cache.tiposDocumento;
  }

  getSubtiposByTipo(tipoId: string): Array<{ id: string; title: string }> {
    if (!this.cache) return [];

    return this.cache.subtipos
      .filter(subtipo => subtipo.tipoId === tipoId)
      .sort((a, b) => a.title.localeCompare(b.title));
  }

  getCacheMetadata(): CacheMetadata | null {
    if (!this.cache) return null;

    const lastUpdated = new Date(this.cache.lastUpdated);
    const recordCount = this.cache.clientes.length +
      this.cache.asuntos.length +
      this.cache.subasuntos.length;

    return {
      lastUpdated,
      version: this.cache.version,
      recordCount,
      isStale: this.isStale(this.cache)
    };
  }

  async forceUpdate(): Promise<CacheData> {
    this.clearCache();
    return await this.initialize();
  }

  private loadFromStorage(): CacheData | null {
    try {
      const data = localStorage.getItem(this.CACHE_KEY);
      return data ? JSON.parse(data) as CacheData : null;
    } catch (error) {
      return null;
    }
  }

  private saveToStorage(data: CacheData): void {
    try {
      localStorage.setItem(this.CACHE_KEY, JSON.stringify(data));
    } catch (error) {
      console.error('Error guardando cache:', error);
    }
  }

  private isValidCache(data: CacheData): boolean {
    if (!data || !data.version) return false;
    return data.version === this.CACHE_VERSION;
  }

  private isStale(data: CacheData): boolean {
    const lastUpdated = new Date(data.lastUpdated);
    const now = new Date();
    const hoursDiff = (now.getTime() - lastUpdated.getTime()) / (1000 * 60 * 60);
    return hoursDiff > this.STALE_THRESHOLD_HOURS;
  }

  private normalizeText(text: string): string {
    return text.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  }

  clearCache(): void {
    try {
      localStorage.removeItem(this.CACHE_KEY);
      this.cache = null;
      this.loadingPromises.clear();
    } catch (error) {
      console.error('Error limpiando cache:', error);
    }
  }

  isCurrentlyUpdating(): boolean {
    return this.isUpdating;
  }
}