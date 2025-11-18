// sharepoint-data.service.ts
import { CommonProps } from "../interfaces/interfaces";
import { CacheManager } from "./cache-manager";

export class SharePointDataService {
  private accessToken: string;
  private cacheManager: CacheManager;
  private isInitialized = false;
  private originalDataService: any;
  /* private readonly SITE_URL = "https://hughesandhughesuy.sharepoint.com/sites/GestorDocumental";
   private readonly API_BASE = "https://graph.microsoft.com/v1.0";*/

  // IDs de las listas de SharePoint
  private readonly LIST_IDS = {
    clientes: "d27e0216-23fb-405a-a567-677561e21701",
    asuntos: "002e0cc5-ba7e-4975-837a-3ee66f163646",
    subasuntos: "7b20956b-1a07-482d-8516-2b6d24fd22f5",
    tiposDocumento: "68ef3918-4451-4586-a2c3-e93af01b2a4e",
    subTiposdocumento: "a90d4a9e-5c48-4f9b-8e82-7c3dc31f19b4",
    CARPETA: "2f3609c6-afd3-41a7-a1ea-91c103986dd5",
    NIVEL1: "44f371bb-5ca6-43c2-ad9c-281d612e1f92",
    NIVEL2: "612c8c64-15c3-4cf4-aa20-5205dffc0536",
    Tema: "fdfbc633-01ea-4efe-81a0-accbdf4bd77f",
    Subtema: "2ed1c934-c215-479d-b053-b30ef133e8c9",
    TipoDoc: "b06f75b7-9f96-42ee-9c91-c5b1c7e808f3",
    CARPETA1_DJ: "f28c04ec-c279-4f1f-bfe9-32aacce00ecc",
    CARPETA1: "cffdd944-add5-4314-bc9c-a40e5c1785f1",
    CARPETA2: "0a66e16a-a6ee-4e99-811d-da00c920f80d",
    CARPETA3: "a84a5331-07cb-42b9-993b-c709167b58c7",
    CARPETA4: "78457760-394f-4786-a729-db6c8e839ccb",
    CARPETA5: "9b53ef21-3e8b-4ca4-a037-85f9e362d118",
    CARPETA6: null,
    CARPETA7: null,
  };

  private cache = new Map<string, any>();

  constructor(accessToken: string) {
    this.accessToken = accessToken;
    this.originalDataService = this;
    this.cacheManager = new CacheManager(this.originalDataService);
  }

  async initialize(): Promise<void> {
    if (this.isInitialized) return;

    try {
      await this.cacheManager.initialize();
      this.isInitialized = true;
    } catch (error) {
      console.error('Error inicializando SharePointDataService:', error);
      // Marcar como inicializado para que funcione sin cache
      this.isInitialized = true;
    }
  }

  private async getSiteId(): Promise<string> {
    const cacheKey = 'siteId';
    const cached = this.cache.get(cacheKey);
    if (cached) return cached.data;

    try {
      const response = await fetch(
        "https://graph.microsoft.com/v1.0/sites/hughesandhughesuy.sharepoint.com:/sites/GestorDocumental",
        {
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Accept': 'application/json'
          }
        }
      );

      if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);

      const data = await response.json();
      this.cache.set(cacheKey, { data: data.id, timestamp: Date.now() });
      return data.id;
    } catch (error) {
      console.error("Error getting site ID:", error);
      throw error;
    }
  }

  private async fetchPaginatedData<T>(url: string, mapper: (item: any) => T): Promise<T[]> {
    let allItems: T[] = [];
    let nextLink: string | null = url;

    while (nextLink && allItems.length < 10000) {
      try {
        const response = await fetch(nextLink, {
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Accept': 'application/json'
          }
        });

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const data = await response.json();
        allItems = [...allItems, ...data.value.map(mapper)];
        nextLink = data["@odata.nextLink"] || null;
      } catch (error) {
        console.error("Error in paginated fetch:", error);
        throw error;
      }
    }

    return allItems;
  }

  private async fetchFromSharePoint<T>(listId: string, mapper: (item: any) => T, filterFn?: (item: T) => boolean): Promise<T[]> {
    if (!listId) return [];

    try {
      const siteId = await this.getSiteId();
      const apiUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=1000`;
      const allItems = await this.fetchPaginatedData<T>(apiUrl, mapper);
      return filterFn ? allItems.filter(filterFn) : allItems;

    } catch (error) {
      console.error(`[ERROR en fetchFromSharePoint]`, {
        message: error.message,
        listId
      });
      return [];
    }
  }

  // CLIENTES
  async getClientes(searchText: string = '', maxResults: number = 100): Promise<Array<{ id: string; title: string }>> {
    await this.ensureInitialized();

    try {
      // Intentar desde cache primero
      const cacheResults = this.cacheManager.getClientes(searchText, maxResults);
      if (cacheResults.length > 0) {
        return cacheResults;
      }
      return await this.getClientesOriginal();
    } catch (error) {
      console.error('Error en getClientes:', error);
      return [];
    }
  }
  async getAsuntosByCliente(clienteId: string, clienteTitle?: string): Promise<Array<{ id: string; title: string }>> {
    await this.ensureInitialized();

    try {
      const cacheResults = await this.cacheManager.getAsuntosByCliente(clienteId);
      if (cacheResults.length > 0) {
        return cacheResults;
      }
      return await this.getAsuntosByClienteOriginal(clienteId, clienteTitle);
    } catch (error) {
      console.error('Error en getAsuntosByCliente:', error);
      return [];
    }
  }

  async getSubasuntosByAsunto(asuntoId: string, asuntoTitle?: string): Promise<Array<{ id: string; title: string }>> {
    await this.ensureInitialized();

    try {
      const cacheResults = this.cacheManager.getSubasuntosByAsunto(asuntoId);
      if (cacheResults.length > 0) {
        return cacheResults;
      }
      return await this.getSubasuntosByAsuntoOriginal(asuntoId, asuntoTitle);
    } catch (error) {
      console.error('Error en getSubasuntosByAsunto:', error);
      return [];
    }
  }

  getCacheMetadata() {
    return this.cacheManager.getCacheMetadata();
  }

  async forceRefreshCache(): Promise<void> {
    await this.cacheManager.forceUpdate();
  }

  clearCache(): void {
    this.cacheManager.clearCache();
  }

  isUpdatingCache(): boolean {
    return this.cacheManager.isCurrentlyUpdating();
  }

  // MÉTODOS PRIVADOS
  private async ensureInitialized(): Promise<void> {
    if (!this.isInitialized) {
      await this.initialize();
    }
  }

  // Método para actualizar el token de acceso
  updateAccessToken(newToken: string): void {
    this.accessToken = newToken;
    // También actualizar en la instancia original si es necesario
  }

  private async getClientesOriginal(): Promise<Array<{ id: string; title: string }>> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.clientes,
      item => ({
        id: item.fields?.field_0?.toString()?.trim() || item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.Title?.toString()?.trim() || 'Sin título'
      })
    );
  }

  private async getAsuntosByClienteOriginal(clienteId: string, clienteTitle?: string): Promise<Array<{ id: string; title: string }>> {
    return await this.fetchFromSharePoint<any>(
      this.LIST_IDS.asuntos,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
        clienteId: item.fields?.field_2?.toString()?.trim() || '',
        clienteName: item.fields?.field_3?.toString()?.trim() || '',
        rawFields: {
          Title: item.fields?.Title,
          field_1: item.fields?.field_1,
          field_2: item.fields?.field_2,
          field_3: item.fields?.field_3
        }
      }),
      item => {
        // Include asuntos that match the specific cliente
        const matchesSpecificCliente = item.clienteId?.toString().trim() === clienteId.trim() && 
                                      item.clienteName?.toString().trim().toLowerCase() === clienteTitle.toLowerCase().trim();
        
        // Include asuntos that have null cliente (should appear for all clientes)
        const hasNullCliente = (!item.clienteId || item.clienteId.toString().trim() === '') && 
                              (!item.clienteName || item.clienteName.toString().trim() === '');
        
        return matchesSpecificCliente || hasNullCliente;
      }
    );
  }

  public async getSubasuntosByAsuntoOriginal(asuntoId: string, asuntoTitle?: string): Promise<Array<{ id: string; title: string }>> {
    return await this.fetchFromSharePoint<any>(
      this.LIST_IDS.subasuntos,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
        asuntoId: item.fields?.field_3?.toString()?.trim() || '',
        asuntoTitle: item.fields?.field_4?.toString()?.trim() || '',
        rawFields: item.fields
      }),
      item => item.asuntoId?.toString().trim() === asuntoId.trim() && item.asuntoTitle?.toString().trim().toLowerCase() === asuntoTitle.toLowerCase().trim()
    );
  }

  async getTiposDocumento(): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.tiposDocumento,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título'
      })
    );
  }

  async getSubTiposDocumento(tipoDocumentoId: string): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<any>(
      this.LIST_IDS.subTiposdocumento,
      item => ({
        id: item.id,
        title: item.fields?.field_3 || 'Sin título',
        tipoDocId: item.fields?.field_0 || '',
        description: item.fields?.field_3 || ''
      }),
      item => item.tipoDocId === tipoDocumentoId
    );
  };

  // ADMINISTRACIÓN RRHH
  async getCarpetaRRHH(): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.CARPETA,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.CARPETADESCRIPTION?.toString()?.trim() || item.fields?.Title?.toString()?.trim() || 'Sin título'
      })
    );
  }

  // CONSULADO AUSTRALIA
  async getNivel1(): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.NIVEL1,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.NIVEL1DESCRIPTION?.toString()?.trim() || item.fields?.Title?.toString()?.trim() || 'Sin título'
      })
    );
  }

  async getNivel2(): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.NIVEL2,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.Title?.toString()?.trim() || 'Sin título'
      })
    );
  }

  // CONTADURÍA
  async getTemasContaduria(): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.Tema,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título'
      })
    );
  }

  async getSubtemasByTema(temaId: string): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<any>(
      this.LIST_IDS.Subtema,
      item => {
        const mappedItem = {
          id: item.fields?.Title?.toString()?.trim() || 'Sin código',
          title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
          temaId: item.fields?.field_2?.toString()?.trim() || '',
          rawFields: {
            Title: item.fields?.Title,
            field_1: item.fields?.field_1,
            field_2: item.fields?.field_2
          }
        };
        return mappedItem;
      },
      item => {
        const temaIdMatch = item.temaId === temaId;
        const temaIdPartialMatch = item.temaId.toLowerCase().includes(temaId.toLowerCase());
        const temaIdAsNameMatch = item.temaId.toLowerCase().trim() === temaId.toLowerCase().trim();
        const match = temaIdMatch || temaIdPartialMatch || temaIdAsNameMatch;
        return match;
      }
    );
  }

  async getTipoDocContaduria(): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.TipoDoc,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título'
      })
    );
  }

  async getCarpetaDJ(): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.CARPETA1_DJ,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.Title?.toString()?.trim() || 'Sin título'
      })
    );
  }

  async getCarpeta1(): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<CommonProps>(
      this.LIST_IDS.CARPETA1,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título'
      })
    );
  }

  async getCarpeta2ByCarpeta1(carpeta1Id: string): Promise<CommonProps[]> {
    return await this.fetchFromSharePoint<any>(
      this.LIST_IDS.CARPETA2,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
        carpetaId: item.fields?.field_2?.toString()?.trim() || ''
      }),
      item => item.carpetaId === carpeta1Id.trim()
    );
  }

  async getCarpeta3ByCarpeta2(carpeta2Id: string, carpeta2Title?: string): Promise<CommonProps[]> {
    const normalizedCarpeta2Id = carpeta2Id?.toString()?.trim() || '';
    const normalizedCarpeta2Title = carpeta2Title?.toString()?.trim().toLowerCase() || '';
    
    // Si no tenemos el título, obtenerlo desde Carpeta2
    let actualCarpeta2Title = normalizedCarpeta2Title;
    if (!actualCarpeta2Title && normalizedCarpeta2Id) {
      try {
        const carpeta2Items = await this.fetchFromSharePoint<any>(
          this.LIST_IDS.CARPETA2,
          item => ({
            id: item.fields?.Title?.toString()?.trim() || '',
            title: item.fields?.field_1?.toString()?.trim() || ''
          }),
          item => item.id === normalizedCarpeta2Id
        );
        if (carpeta2Items.length > 0) {
          actualCarpeta2Title = carpeta2Items[0].title?.toString()?.trim().toLowerCase() || '';
        }
      } catch (error) {
        // Silently fail
      }
    }
    
    const allItems = await this.fetchFromSharePoint<any>(
      this.LIST_IDS.CARPETA3,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
        carpetaId: item.fields?.field_2?.toString()?.trim() || '',
        carpetaTitleField3: item.fields?.field_3?.toString()?.trim() || '',
        carpetaTitleField4: item.fields?.field_4?.toString()?.trim() || '',
        carpetaTitleField5: item.fields?.field_5?.toString()?.trim() || '',
        allFields: item.fields
      })
    );
    
    const filteredItems = allItems.filter(item => {
      const normalizedItemCarpetaId = item.carpetaId?.toString()?.trim() || '';
      
      if (normalizedItemCarpetaId !== normalizedCarpeta2Id) {
        return false;
      }
      
      if (actualCarpeta2Title) {
        const field3Match = item.carpetaTitleField3?.toString()?.trim().toLowerCase() === actualCarpeta2Title;
        const field4Match = item.carpetaTitleField4?.toString()?.trim().toLowerCase() === actualCarpeta2Title;
        const field5Match = item.carpetaTitleField5?.toString()?.trim().toLowerCase() === actualCarpeta2Title;
        
        let foundInAnyField = false;
        if (item.allFields) {
          for (const [, value] of Object.entries(item.allFields)) {
            if (value && value.toString().trim().toLowerCase() === actualCarpeta2Title) {
              foundInAnyField = true;
              break;
            }
          }
        }
        
        return field3Match || field4Match || field5Match || foundInAnyField;
      }
      
      return true;
    });
    
    return filteredItems.map(item => ({
      id: item.id,
      title: item.title
    }));
  }

  async getCarpeta4ByCarpeta3(carpeta3Id: string, carpeta3Title?: string): Promise<CommonProps[]> {
    const normalizedCarpeta3Id = carpeta3Id?.toString()?.trim() || '';
    const normalizedCarpeta3Title = carpeta3Title?.toString()?.trim().toLowerCase() || '';
    
    // Si no tenemos el título, obtenerlo desde Carpeta3
    let actualCarpeta3Title = normalizedCarpeta3Title;
    if (!actualCarpeta3Title && normalizedCarpeta3Id) {
      try {
        const carpeta3Items = await this.fetchFromSharePoint<any>(
          this.LIST_IDS.CARPETA3,
          item => ({
            id: item.fields?.Title?.toString()?.trim() || '',
            title: item.fields?.field_1?.toString()?.trim() || ''
          }),
          item => item.id === normalizedCarpeta3Id
        );
        if (carpeta3Items.length > 0) {
          actualCarpeta3Title = carpeta3Items[0].title?.toString()?.trim().toLowerCase() || '';
        }
      } catch (error) {
        // Silently fail
      }
    }
    
    const allItems = await this.fetchFromSharePoint<any>(
      this.LIST_IDS.CARPETA4,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
        carpetaId: item.fields?.field_2?.toString()?.trim() || '',
        carpetaTitleField3: item.fields?.field_3?.toString()?.trim() || '',
        carpetaTitleField4: item.fields?.field_4?.toString()?.trim() || '',
        carpetaTitleField5: item.fields?.field_5?.toString()?.trim() || '',
        allFields: item.fields
      })
    );
    
    const filteredItems = allItems.filter(item => {
      const normalizedItemCarpetaId = item.carpetaId?.toString()?.trim() || '';
      
      if (normalizedItemCarpetaId !== normalizedCarpeta3Id) {
        return false;
      }
      
      if (actualCarpeta3Title) {
        const field3Match = item.carpetaTitleField3?.toString()?.trim().toLowerCase() === actualCarpeta3Title;
        const field4Match = item.carpetaTitleField4?.toString()?.trim().toLowerCase() === actualCarpeta3Title;
        const field5Match = item.carpetaTitleField5?.toString()?.trim().toLowerCase() === actualCarpeta3Title;
        
        let foundInAnyField = false;
        if (item.allFields) {
          for (const [, value] of Object.entries(item.allFields)) {
            if (value && value.toString().trim().toLowerCase() === actualCarpeta3Title) {
              foundInAnyField = true;
              break;
            }
          }
        }
        
        return field3Match || field4Match || field5Match || foundInAnyField;
      }
      
      return true;
    });
    
    return filteredItems.map(item => ({
      id: item.id,
      title: item.title
    }));
  }

  async getCarpeta5ByCarpeta4(carpeta4Id: string, carpeta4Title?: string): Promise<CommonProps[]> {
    const normalizedCarpeta4Id = carpeta4Id?.toString()?.trim() || '';
    const normalizedCarpeta4Title = carpeta4Title?.toString()?.trim().toLowerCase() || '';
    
    // Si no tenemos el título, obtenerlo desde Carpeta4
    let actualCarpeta4Title = normalizedCarpeta4Title;
    if (!actualCarpeta4Title && normalizedCarpeta4Id) {
      try {
        const carpeta4Items = await this.fetchFromSharePoint<any>(
          this.LIST_IDS.CARPETA4,
          item => ({
            id: item.fields?.Title?.toString()?.trim() || '',
            title: item.fields?.field_1?.toString()?.trim() || ''
          }),
          item => item.id === normalizedCarpeta4Id
        );
        if (carpeta4Items.length > 0) {
          actualCarpeta4Title = carpeta4Items[0].title?.toString()?.trim().toLowerCase() || '';
        }
      } catch (error) {
        // Silently fail
      }
    }
    
    const allItems = await this.fetchFromSharePoint<any>(
      this.LIST_IDS.CARPETA5,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
        carpetaId: item.fields?.field_2?.toString()?.trim() || '',
        carpetaTitleField3: item.fields?.field_3?.toString()?.trim() || '',
        carpetaTitleField4: item.fields?.field_4?.toString()?.trim() || '',
        carpetaTitleField5: item.fields?.field_5?.toString()?.trim() || '',
        allFields: item.fields
      })
    );
    
    const filteredItems = allItems.filter(item => {
      const normalizedItemCarpetaId = item.carpetaId?.toString()?.trim() || '';
      
      if (normalizedItemCarpetaId !== normalizedCarpeta4Id) {
        return false;
      }
      
      if (actualCarpeta4Title) {
        const field3Match = item.carpetaTitleField3?.toString()?.trim().toLowerCase() === actualCarpeta4Title;
        const field4Match = item.carpetaTitleField4?.toString()?.trim().toLowerCase() === actualCarpeta4Title;
        const field5Match = item.carpetaTitleField5?.toString()?.trim().toLowerCase() === actualCarpeta4Title;
        
        let foundInAnyField = false;
        if (item.allFields) {
          for (const [, value] of Object.entries(item.allFields)) {
            if (value && value.toString().trim().toLowerCase() === actualCarpeta4Title) {
              foundInAnyField = true;
              break;
            }
          }
        }
        
        return field3Match || field4Match || field5Match || foundInAnyField;
      }
      
      return true;
    });
    
    return filteredItems.map(item => ({
      id: item.id,
      title: item.title
    }));
  }

  async getCarpeta6ByCarpeta5(carpeta5Id: string): Promise<CommonProps[]> {
    if (!this.LIST_IDS.CARPETA6) {
      console.warn('CARPETA6 ID is null');
      return [];
    }
    return await this.fetchFromSharePoint<any>(
      this.LIST_IDS.CARPETA6,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
        carpetaId: item.fields?.field_2?.toString()?.trim() || ''
      }),
      item => item.carpetaId === carpeta5Id.trim()
    );
  }

  async getCarpeta7ByCarpeta6(carpeta6Id: string): Promise<CommonProps[]> {
    if (!this.LIST_IDS.CARPETA7) {
      console.warn('CARPETA7 ID is null');
      return [];
    }
    return await this.fetchFromSharePoint<any>(
      this.LIST_IDS.CARPETA7,
      item => ({
        id: item.fields?.Title?.toString()?.trim() || 'Sin código',
        title: item.fields?.field_1?.toString()?.trim() || 'Sin título',
        carpetaId: item.fields?.field_2?.toString()?.trim() || ''
      }),
      item => item.carpetaId === carpeta6Id.trim()
    );
  }

  async searchClientes(searchTerm: string, pageSize: number = 100): Promise<{
    items: CommonProps[],
    hasMore: boolean,
    nextSkipToken?: string
  }> {
    try {
      const allClientes = await this.getClientes();
      const normalizedSearch = searchTerm.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

      const filteredClientes = allClientes.filter(cliente =>
        cliente.title.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes(normalizedSearch) ||
        cliente.id.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes(normalizedSearch)
      ).slice(0, pageSize);

      return {
        items: filteredClientes,
        hasMore: false,
        nextSkipToken: undefined
      };

    } catch (error) {
      console.error('Error en searchClientes:', error);
      return { items: [], hasMore: false };
    }
  }

  getCacheStats(): { size: number, keys: string[] } {
    return {
      size: this.cache.size,
      keys: Array.from(this.cache.keys())
    };
  }
}