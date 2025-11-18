import * as React from "react";
import { TextField, PrimaryButton, MessageBar, MessageBarType, Spinner, SpinnerSize, Dropdown, IDropdownOption, DialogFooter, DefaultButton, Dialog, DialogType } from "office-ui-fabric-react";
import { findDocumentLibrary, getSharePointDrives, getSiteId, uploadToSharePoint } from "../../utilities/sharepoint_helpers";
import { BIBLIOTECAS_DISPONIBLES } from "../constants/sharepointConfig";
import { SharePointDataService } from "../services/sharepoint-data.service";
import { MetadataComponents } from "./metadata-components";
import { DocumentTitleEditorProps, DocumentTitleEditorState, SharePointResponse } from "../interfaces/interfaces";

export default class DocumentTitleEditor extends React.Component<DocumentTitleEditorProps, DocumentTitleEditorState> {
  private dataService: SharePointDataService;
  private readonly SITE_URL = "https://hughesandhughesuy.sharepoint.com/sites/GestorDocumental";
  private readonly MAX_RETRY_ATTEMPTS = 3;
  private readonly METADATA_QUEUE_LIST_ID = "d81f68e5-fdb6-4555-9739-864878646ae8";
  private isDataServiceInitialized = false;

  constructor(props: DocumentTitleEditorProps) {
    super(props);
    this.state = {
      currentTitle: "",
      newTitle: "",
      baseTitle: "",
      isLoading: true,
      isSaving: false,
      message: "",
      messageType: MessageBarType.info,
      showMessage: false,
      officeApp: "",
      selectedBiblioteca: "",
      bibliotecas: BIBLIOTECAS_DISPONIBLES,
      currentMetadata: {},
      existingMetadata: null,
      isExistingDocument: false,
      currentFileId: "",
      currentDocId: "",
      isLoadingExistingMetadata: false,
      originalDocumentSaved: false,
      isInCellEditMode: false,
      preservedComponentState: null,

      pendingSync: false,
      localChanges: {},
      lastSyncAttempt: null,
      syncRetryCount: 0,

      showCloseDocumentDialog: false,
      pendingSaveData: undefined,
      isSyncing: false,
    };

    this.metadataComponentRef = React.createRef();
    this.dataService = new SharePointDataService(props.accessToken);
  }

  private initializeDataService = async () => {
    if (this.isDataServiceInitialized) {
      return;
    }
    try {
      if (this.state.selectedBiblioteca === "DOCUMENTOS_CLIENTES") {
        await this.dataService.initialize();
        const cacheMetadata = this.dataService.getCacheMetadata();
        if (cacheMetadata && cacheMetadata.recordCount > 0) {
          if (cacheMetadata.isStale) {
            this.setState({
              message: "Los datos están siendo actualizados en segundo plano para garantizar la información más reciente.",
              messageType: MessageBarType.info,
              showMessage: true
            });
          } else if (cacheMetadata.recordCount > 0) {
            this.setState({
              message: `Cache activo: ${cacheMetadata.recordCount.toLocaleString()} registros listos para búsqueda instantánea.`,
              messageType: MessageBarType.success,
              showMessage: true
            });
          }
        }
      }
    } catch (error) {
      this.isDataServiceInitialized = false;
      this.setState({
        message: "Error inicializando datos. Algunos metadatos podrían cargar más lento de lo normal.",
        messageType: MessageBarType.warning,
        showMessage: true
      });
    }
  };


  private metadataComponentRef: React.RefObject<MetadataComponents>;
  async componentDidMount() {
    await this.loadCurrentTitle();
    await this.detectCurrentLibrary();
    await this.initializeDataService();
    
    // Sincronización automática al cargar el complemento
    setTimeout(async () => {
      await this.handleForceSync(false); // No mostrar mensajes en sincronización automática
    }, 2000); // Esperar 2 segundos para que termine la carga inicial
    
    if (this.state.pendingSync) {
      setTimeout(() => {
        this.attemptBackgroundSync();
      }, 4000);
    }
  }

  //   private renderCacheControls = () => {
  //   const cacheMetadata = this.dataService.getCacheMetadata();
  //   const isUpdating = this.dataService.isUpdatingCache();

  //   if (!cacheMetadata) return null;

  //   return (
  //     <div style={{ 
  //       padding: '10px', 
  //       backgroundColor: '#f8f9fa', 
  //       border: '1px solid #e0e0e0', 
  //       borderRadius: '4px',
  //       marginBottom: '16px',
  //       fontSize: '12px'
  //     }}>
  //       <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
  //         <span>
  //           📊 Cache: {cacheMetadata.recordCount.toLocaleString()} registros
  //           {cacheMetadata.isStale && <span style={{ color: '#ff8c00' }}> (actualizando...)</span>}
  //         </span>
  //         <div>
  //           <button
  //             onClick={() => this.dataService.forceRefreshCache()}
  //             disabled={isUpdating}
  //             style={{
  //               padding: '4px 8px',
  //               fontSize: '11px',
  //               marginRight: '8px',
  //               border: '1px solid #ccc',
  //               borderRadius: '3px',
  //               backgroundColor: isUpdating ? '#f0f0f0' : 'white',
  //               cursor: isUpdating ? 'not-allowed' : 'pointer'
  //             }}
  //           >
  //             {isUpdating ? '🔄 Actualizando...' : '🔄 Actualizar'}
  //           </button>
  //           <button
  //             onClick={() => {
  //               this.dataService.clearCache();
  //               this.setState({
  //                 message: "Cache limpiado. Los datos se recargarán en la próxima sesión.",
  //                 messageType: MessageBarType.info,
  //                 showMessage: true
  //               });
  //             }}
  //             style={{
  //               padding: '4px 8px',
  //               fontSize: '11px',
  //               border: '1px solid #ccc',
  //               borderRadius: '3px',
  //               backgroundColor: 'white',
  //               cursor: 'pointer'
  //             }}
  //           >
  //             🗑️ Limpiar
  //           </button>
  //         </div>
  //       </div>
  //       <div style={{ marginTop: '4px', color: '#666' }}>
  //         Última actualización: {cacheMetadata.lastUpdated.toLocaleString()}
  //       </div>
  //     </div>
  //   );
  // };

  componentDidUpdate(prevProps: DocumentTitleEditorProps) {
    if (prevProps.accessToken !== this.props.accessToken) {
      this.dataService = new SharePointDataService(this.props.accessToken);
      this.isDataServiceInitialized = false;
    }
  }

  private exitCellEditMode = async (): Promise<boolean> => {
    if (this.state.officeApp !== "Excel") return true;
    try {
      await Excel.run(async (context) => {
        const application = context.workbook.application;
        try {
          application.calculate(Excel.CalculationType.recalculate);
          await context.sync();
          const activeCell = context.workbook.getActiveCell();
          activeCell.load("address");
          await context.sync();
          return true;
        } catch (editError) {
          console.warn("Intento de salir del modo de edición:", editError);
          throw editError;
        }
      });

      this.setState({ isInCellEditMode: false });
      return true;
    } catch (error) {
      console.error("No se pudo salir del modo de edición:", error);
      this.setState({
        isInCellEditMode: true,
        message: "Excel está en modo de edición de celdas. Por favor, presione Enter o Tab para salir del modo de edición y vuelva a intentar.",
        messageType: MessageBarType.warning,
        showMessage: true
      });
      return false;
    }
  };

  loadCurrentTitle = async () => {
    try {
      this.setState({ isLoading: true, isInCellEditMode: false });
      const officeApp = this.detectOfficeApplication();
      this.setState({ officeApp });
      if (officeApp === "Excel") {
        const canProceed = await this.exitCellEditMode();
        if (!canProceed) {
          this.setState({ isLoading: false });
          return;
        }
      }

      let title = "";
      let isDocumentSaved = false;
      if (officeApp === "Excel") {
        await Excel.run(async (context) => {
          const workbook = context.workbook;
          workbook.load("name");
          const properties = workbook.properties;
          properties.load("title");
          workbook.load("isDirty");
          await context.sync();
          // title = properties.title || workbook.name || "Untitled Workbook";
          title = workbook.name || "Untitled Workbook";
          isDocumentSaved = !workbook.isDirty;
        });

      } else if (officeApp === "Word") {
        await Word.run(async (context) => {
          const document = context.document;
          document.load("saved");
          const properties = document.properties;
          console.log(properties);
          properties.load("title");
          await context.sync();
          title = properties.title || "Untitled Document";
          isDocumentSaved = document.saved;
        });
      } else if (officeApp === "PowerPoint") {
        await PowerPoint.run(async (context) => {
          const presentation = context.presentation;
          presentation.load("title");
          await context.sync();
          title = presentation.title || "Untitled Presentation";
          isDocumentSaved = true;
        });
      } else {
        title = "Document";
      }

      const baseTitle = this.extractBaseTitle(title);
      this.setState({
        currentTitle: title,
        newTitle: baseTitle,
        baseTitle: baseTitle,
        isLoading: false,
        originalDocumentSaved: isDocumentSaved
      });
      await this.loadAvailableLibraries();
    } catch (error) {
      console.error("Error loading document title:", error);
      if (error.message && error.message.includes("InvalidOperationInCellEditMode")) {
        this.setState({
          isLoading: false,
          isInCellEditMode: true,
          showMessage: true,
          currentTitle: "Modo edición activo",
          newTitle: ""
        });
      } else {
        this.setState({
          isLoading: false,
          message: "Error loading document title: " + error.toString(),
          messageType: MessageBarType.error,
          showMessage: true,
          currentTitle: "Error loading title",
          newTitle: ""
        });
      }
    }
  };

  private extractBaseTitle = (fullTitle: string): string => {
    let titleWithoutExtension = fullTitle;
    const extensions = ['.xlsx', '.docx', '.pptx', '.xls', '.doc', '.ppt'];
    
    for (const ext of extensions) {
      if (titleWithoutExtension.toLowerCase().endsWith(ext.toLowerCase())) {
        titleWithoutExtension = titleWithoutExtension.substring(0, titleWithoutExtension.length - ext.length);
        break;
      }
    }
    const docIdPattern = /-\d+$/;
    const baseTitle = titleWithoutExtension.replace(docIdPattern, '');
    
    return baseTitle;
  };

  // private extractBaseTitle = (fullTitle: string): string => {
  //   const fileExtension = this.getFileExtension();
  //   let titleWithoutExtension = fullTitle.replace(fileExtension, '');
  //   const docIdPattern = /-\d+$/;
  //   const baseTitle = titleWithoutExtension.replace(docIdPattern, '');
  //   return baseTitle;
  // };

  private extractDocId = (fullTitle: string): string => {
    const fileExtension = this.getFileExtension();
    const titleWithoutExtension = fullTitle.replace(fileExtension, '');
    const docIdMatch = titleWithoutExtension.match(/-(\d+)$/);
    return docIdMatch ? docIdMatch[1] : "";
  };

  detectOfficeApplication = (): string => {
    if (typeof Excel !== "undefined") {
      return "Excel";
    } else if (typeof Word !== "undefined") {
      return "Word";
    } else if (typeof PowerPoint !== "undefined") {
      return "PowerPoint";
    }
    return "Unknown";
  };

  updateTitle = async () => {
    if (!this.state.newTitle.trim()) {
      this.setState({
        message: "Please enter a valid title",
        messageType: MessageBarType.warning,
        showMessage: true
      });
      return;
    }

    if (this.state.officeApp === "Excel") {
      const canProceed = await this.exitCellEditMode();
      if (!canProceed) return;
    }

    try {
      // this.setState({ isLoading: true });
      let fullTitle = this.state.newTitle.trim();
      if (this.state.currentDocId) { fullTitle = `${fullTitle}-${this.state.currentDocId}`; }
      if (this.state.officeApp === "Excel") {
        await Excel.run(async (context) => {
          const workbook: any = context.workbook;
          const properties = workbook.properties;
          properties.title = fullTitle;
          try {
            workbook.name = fullTitle + this.getFileExtension();
          } catch (nameError) {
            console.warn("No se pudo cambiar el nombre del workbook:", nameError);
          }

          await context.sync();
        });
      } else if (this.state.officeApp === "Word") {
        await Word.run(async (context) => {
          const properties = context.document.properties;
          properties.title = fullTitle;
          properties.subject = fullTitle;
          await context.sync();
        });

      } else if (this.state.officeApp === "PowerPoint") {
        await PowerPoint.run(async (context) => {
          const presentation: any = context.presentation;
          presentation.title = fullTitle;
          await context.sync();
        });
      }

      this.setState({
        currentTitle: fullTitle,
        baseTitle: this.state.newTitle.trim(),
        isLoading: false,
        message: `El documento se actualizó correctamente en ${this.state.officeApp}!`,
        messageType: MessageBarType.success,
        showMessage: false
      });

    } catch (error) {
      console.error("Error updating document title:", error);
      this.setState({
        isLoading: false,
        message: "Error updating title: " + error.toString(),
        messageType: MessageBarType.error,
        showMessage: true
      });
    }
  };

  handleTitleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ newTitle: newValue || "" });
  };

  handleBibliotecaChange = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const bibliotecaId = option?.key as string || "";
    this.setState({
      selectedBiblioteca: bibliotecaId,
      currentMetadata: {},
      existingMetadata: null
    });

    if (bibliotecaId === "DOCUMENTOS_CLIENTES" && !this.isDataServiceInitialized) {
      await this.initializeDataService();
    }

    if (this.state.isExistingDocument && bibliotecaId) {
      await this.loadExistingDocumentMetadata(bibliotecaId);
    }
  };

  handleMetadataChange = (metadata: any) => {
    this.setState(prevState => ({
      currentMetadata: {
        ...prevState.currentMetadata,
        ...metadata
      }
    }));

  }

  detectCurrentLibrary = async () => {
    try {
      if (!this.state.currentTitle || this.state.currentTitle === "Untitled Document" ||
        this.state.currentTitle === "Untitled Workbook" || this.state.currentTitle === "Untitled Presentation") {
        this.setState({
          isExistingDocument: false,
          message: "Documento nuevo - seleccione una biblioteca para continuar.",
          messageType: MessageBarType.info,
          showMessage: true
        });
        return;
      }
      let documentUrl = "";
      try {
        if (this.state.officeApp === "Excel") {
          await Excel.run(async (context) => {
            const workbook: any = context.workbook;
            workbook.load("path");
            await context.sync();
            documentUrl = workbook.path || "";
          });
        } else if (this.state.officeApp === "Word") {
          await Word.run(async (context) => {
            const document: any = context.document;
            document.load("url");
            await context.sync();
            documentUrl = document.url || "";
          });
        } else if (this.state.officeApp === "PowerPoint") {
          await PowerPoint.run(async (context) => {
            const presentation: any = context.presentation;
            presentation.load("url");
            await context.sync();
            documentUrl = presentation.url || "";
          });
        }
      } catch (urlError) {
        console.log("No se pudo obtener URL del documento:", urlError);
      }

      let bibliotecaFromUrl = this.detectLibraryFromUrl(documentUrl);
      const result: SharePointResponse = await findDocumentLibrary(
        this.props.accessToken,
        this.SITE_URL,
        this.state.currentTitle
      );

      if (result.success && result.data) {
        const driveName = result.data.driveName;
        const fileId = result.data.fileId;
        let matchingLibrary = null;

        if (bibliotecaFromUrl) matchingLibrary = BIBLIOTECAS_DISPONIBLES.find(lib => lib.id === bibliotecaFromUrl);
        if (!matchingLibrary) matchingLibrary = this.findLibraryByDriveName(driveName);
        
        console.log('DEBUG - bibliotecaFromUrl:', bibliotecaFromUrl);
        console.log('DEBUG - driveName:', driveName);
        console.log('DEBUG - matchingLibrary:', matchingLibrary);
        
        if (matchingLibrary) {
          const docId = this.extractDocId(this.state.currentTitle);
          const isExisting = !!docId || !!fileId;
          this.setState({
            selectedBiblioteca: matchingLibrary.id,
            isExistingDocument: isExisting,
            currentFileId: fileId,
            currentDocId: docId,
            originalDocumentSaved: true,
            message: `Documento encontrado en ${matchingLibrary.title}. ${docId ? `DocID: ${docId}` : 'Archivo detectado en SharePoint.'}`,
            messageType: MessageBarType.info,
            showMessage: false
          });
          console.log('DEBUG - calling loadExistingDocumentMetadata with library:', matchingLibrary.id);
          await this.loadExistingDocumentMetadata(matchingLibrary.id);
        } else {
          console.warn('No se pudo determinar la biblioteca correcta para:', driveName);
          this.setState({
            isExistingDocument: !!fileId,
            currentFileId: fileId,
            currentDocId: this.extractDocId(this.state.currentTitle),
            message: `Documento encontrado en SharePoint (Drive: "${driveName}") pero no se pudo determinar la biblioteca automáticamente. Por favor seleccione la biblioteca correcta manualmente.`,
            messageType: MessageBarType.warning,
            showMessage: true
          });
        }
      } else {
        this.setState({
          isExistingDocument: false,
          message: "Documento nuevo - seleccione una biblioteca para continuar.",
          messageType: MessageBarType.info,
          showMessage: true
        });
      }
    } catch (error) {
      this.setState({
        isExistingDocument: false,
        message: "No se pudo detectar si el documento existe en SharePoint. Procederá como documento nuevo.",
        messageType: MessageBarType.warning,
        showMessage: true
      });
    }
  };

  private detectLibraryFromUrl = (url: string): string | null => {
    if (!url) return null;
    
    // Mapeo de URLs/paths a bibliotecas
    const urlMappings = [
      { pattern: /DOCUMENTOS_CLIENTES(_V2)?|DocumentosClientes(_V2)?|Documentos%20Clientes(_V2)?/i, library: "DOCUMENTOS_CLIENTES" },
      { pattern: /DOCUMENTOS_SOCIOS|DocumentosSocios|Documentos%20Socios/i, library: "DOCUMENTOS_SOCIOS" },
      { pattern: /DOCUMENTOS_ADMIN_RRHH|DocumentosAdminRRHH|Documentos%20Admin%20RRHH|AdministracionRRHH/i, library: "DOCUMENTOS_ADMIN_RRHH" },
      { pattern: /DOCUMENTOS_CONSULADO_AUSTRALIA|DocumentosConsuladoAustralia|Documentos%20Consulado%20Australia/i, library: "DOCUMENTOS_CONSULADO_AUSTRALIA" },
      { pattern: /DOCUMENTOS_CONTADURIA|DocumentosContaduria|Documentos%20Contaduria|Contadur[ií]a/i, library: "DOCUMENTOS_CONTADURIA" },
      { pattern: /DOCUMENTOS_DECLARACIONES_JURADAS|DocumentosDeclaracionesJuradas|Documentos%20Declaraciones%20Juradas|DeclaracionesJuradas/i, library: "DOCUMENTOS_DECLARACIONES_JURADAS" },
      { pattern: /DOCUMENTOS_INTERNO|DocumentosInterno|Documentos%20Interno/i, library: "DOCUMENTOS_INTERNO" }
    ];

    for (const mapping of urlMappings) {
      if (mapping.pattern.test(url)) {
        return mapping.library;
      }
    }
    return null;
  };

  private findLibraryByDriveName = (driveName: string): any => {
    if (!driveName || driveName.includes("OneDrive") || driveName.includes("Documents")) return null;
    const driveNameMappings = [
      { patterns: ["DOCUMENTOS CLIENTES", "DocumentosClientes", "Documentos Clientes", "DOCUMENTOS_CLIENTES"], library: "DOCUMENTOS_CLIENTES" },
      { patterns: ["DOCUMENTOS SOCIOS", "DocumentosSocios", "Documentos Socios", "DOCUMENTOS_SOCIOS"], library: "DOCUMENTOS_SOCIOS" },
      { patterns: ["DOCUMENTOS ADMIN RRHH", "DocumentosAdminRRHH", "Documentos Admin RRHH", "DOCUMENTOS_ADMIN_RRHH", "AdministracionRRHH"], library: "DOCUMENTOS_ADMIN_RRHH" },
      { patterns: ["DOCUMENTOS CONSULADO AUSTRALIA", "DocumentosConsuladoAustralia", "Documentos Consulado Australia", "DOCUMENTOS_CONSULADO_AUSTRALIA"], library: "DOCUMENTOS_CONSULADO_AUSTRALIA" },
      { patterns: ["DOCUMENTOS CONTADURIA", "DocumentosContaduria", "Documentos Contaduria", "DOCUMENTOS_CONTADURIA", "Contaduria"], library: "DOCUMENTOS_CONTADURIA" },
      { patterns: ["DOCUMENTOS DECLARACIONES JURADAS", "DocumentosDeclaracionesJuradas", "Documentos Declaraciones Juradas", "DOCUMENTOS_DECLARACIONES_JURADAS", "DeclaracionesJuradas"], library: "DOCUMENTOS_DECLARACIONES_JURADAS" },
      { patterns: ["DOCUMENTOS INTERNO", "DocumentosInterno", "Documentos Interno", "DOCUMENTOS_INTERNO"], library: "DOCUMENTOS_INTERNO" }
    ];

    for (const mapping of driveNameMappings) {
      for (const pattern of mapping.patterns) {
        if (driveName === pattern || driveName.toLowerCase() === pattern.toLowerCase()) {
          return BIBLIOTECAS_DISPONIBLES.find(lib => lib.id === mapping.library);
        }
      }
    }

    for (const mapping of driveNameMappings) {
      for (const pattern of mapping.patterns) {
        if (driveName.toLowerCase().includes(pattern.toLowerCase()) ||
          pattern.toLowerCase().includes(driveName.toLowerCase())) {
          return BIBLIOTECAS_DISPONIBLES.find(lib => lib.id === mapping.library);
        }
      }
    }

    const keywordMappings = [
      { keywords: ["clientes", "client"], library: "DOCUMENTOS_CLIENTES" },
      { keywords: ["socios", "socio", "partner"], library: "DOCUMENTOS_SOCIOS" },
      { keywords: ["admin", "rrhh", "recursos", "humanos"], library: "DOCUMENTOS_ADMIN_RRHH" },
      { keywords: ["consulado", "australia", "embajada"], library: "DOCUMENTOS_CONSULADO_AUSTRALIA" },
      { keywords: ["contaduria", "contable", "accounting"], library: "DOCUMENTOS_CONTADURIA" },
      { keywords: ["declaraciones", "juradas", "dj"], library: "DOCUMENTOS_DECLARACIONES_JURADAS" },
      { keywords: ["interno", "internal"], library: "DOCUMENTOS_INTERNO" }
    ];

    const driveNameLower = driveName.toLowerCase();
    for (const mapping of keywordMappings) {
      for (const keyword of mapping.keywords) {
        if (driveNameLower.includes(keyword)) {
          return BIBLIOTECAS_DISPONIBLES.find(lib => lib.id === mapping.library);
        }
      }
    }
    console.warn(`❌ No se encontró biblioteca para drive: "${driveName}"`);
    return null;
  };

  loadExistingDocumentMetadata = async (bibliotecaId: string) => {
    console.log('DEBUG loadExistingDocumentMetadata - bibliotecaId:', bibliotecaId, 'currentFileId:', this.state.currentFileId);
    if (!this.state.currentFileId) return;
    try {
      this.setState({ isLoadingExistingMetadata: true });
      const siteResponse = await getSiteId(this.props.accessToken, this.SITE_URL);
      if (!siteResponse.success) {
        throw new Error(`Failed to get site: ${siteResponse.error}`);
      }

      const drivesResponse = await getSharePointDrives(this.props.accessToken, siteResponse.data);
      if (!drivesResponse.success) {
        throw new Error(`Failed to get drives: ${drivesResponse.error}`);
      }

      const drives = drivesResponse.data;
      // Use the correct drive name mapping for DOCUMENTOS_CLIENTES
      const driveName = this.getDriveNameFromLibraryId(bibliotecaId);
      console.log('DEBUG loadExistingDocumentMetadata - driveName:', driveName);
      const targetDrive = drives.find((drive: any) =>
        drive.name === driveName ||
        drive.id === driveName
      );
      console.log('DEBUG loadExistingDocumentMetadata - targetDrive found:', targetDrive?.name);

      if (targetDrive) {
        const fileInfoResponse = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${targetDrive.id}/items/${this.state.currentFileId}?$expand=listItem($expand=fields)`,
          {
            headers: {
              Authorization: `Bearer ${this.props.accessToken}`,
            },
          }
        );

        if (fileInfoResponse.ok) {
          const fileData = await fileInfoResponse.json();
          const actualDriveName = fileData.parentReference?.path || "";
          const realLibrary = this.detectLibraryFromUrl(actualDriveName) || this.detectLibraryFromDriveName(fileData.parentReference?.driveId);

          if (realLibrary && realLibrary !== bibliotecaId) {
            this.setState({
              selectedBiblioteca: realLibrary,
              message: `Biblioteca corregida automáticamente a ${BIBLIOTECAS_DISPONIBLES.find(b => b.id === realLibrary)?.title}`,
              messageType: MessageBarType.info,
              showMessage: true
            });

            await this.loadExistingDocumentMetadata(realLibrary);
            return;
          }

          const metadataFields = fileData.listItem?.fields || {};
          if (metadataFields.DocID && !this.state.currentDocId) {
            this.setState({ currentDocId: metadataFields.DocID.toString() });
          }

          const extractLookupValue = (field: any): string => {
            if (!field) return "";
            if (typeof field === 'string') return field;
            if (typeof field === 'object') { return field.LookupValue || field.Label || field.Title || field.toString(); }
            return field.toString();
          };

          const extractChoiceValue = (field: any): string => {
            if (!field) return "";
            if (typeof field === 'string') return field;
            if (Array.isArray(field)) return field.join(", ");
            if (typeof field === 'object') { return field.Value || field.Label || field.toString(); }
            return field.toString();
          };

          const existingMetadata = {
            ...metadataFields,
            DateCreated: metadataFields.DateCreated || metadataFields.Created,
            DateModified: metadataFields.DateModified || metadataFields.Modified,
            DocID: metadataFields.DocID,
            Cliente: extractLookupValue(metadataFields.Cliente),
            Asunto: extractLookupValue(metadataFields.Asunto),
            S_Asunto: extractLookupValue(metadataFields.S_Asunto),
            Tipo_Doc: extractChoiceValue(metadataFields.Tipo_Doc),
            S_Tipo: extractLookupValue(metadataFields.S_Tipo),
            CarpetaRRHH: extractChoiceValue(metadataFields.Carpeta1),
            Nivel1: extractChoiceValue(metadataFields.Nivel1),
            Nivel2: extractChoiceValue(metadataFields.Nivel2),
            Tema: extractLookupValue(metadataFields.Tema),
            SubTema: extractLookupValue(metadataFields.SubTema),
            TipoDoc: extractChoiceValue(metadataFields.TipoDoc),
            CarpetaDJ: extractChoiceValue(metadataFields.Carpeta1),
            Carpeta1: extractLookupValue(metadataFields.Carpeta1),
            Carpeta2: extractLookupValue(metadataFields.Carpeta2),
            Carpeta3: extractLookupValue(metadataFields.Carpeta3),
            Carpeta4: extractLookupValue(metadataFields.Carpeta4),
            Carpeta5: extractLookupValue(metadataFields.Carpeta5),
            Carpeta6: extractLookupValue(metadataFields.Carpeta6),
            Carpeta7: extractLookupValue(metadataFields.Carpeta7),
            ProcesadoAddIn: metadataFields.ProcesadoAddIn,
            autorEmail: metadataFields.autorEmail,
            operadorEmail: metadataFields.operadorEmail,
            ContentType: metadataFields.ContentType,
            FileLeafRef: metadataFields.FileLeafRef,
            FileDirRef: metadataFields.FileDirRef,
            FileRef: metadataFields.FileRef,
            Title: metadataFields.Title,
            Subject: metadataFields.Subject,
            Keywords: metadataFields.Keywords,
            Category: metadataFields.Category,
            Version: metadataFields._UIVersionString || metadataFields.Version,
            CheckoutUser: extractLookupValue(metadataFields.CheckoutUser),
          };
          this.setState({
            existingMetadata: existingMetadata,
            message: "Metadatos existentes cargados correctamente. Los campos se precargarán automáticamente.",
            messageType: MessageBarType.success,
            showMessage: true
          }, () => {
            // Generar metadata inicial
            setTimeout(() => {
              const currentMeta = this.generateCurrentMetadataFromExisting(existingMetadata, bibliotecaId);
              this.setState({
                currentMetadata: currentMeta
              }, () => { console.log("currentMetadata actualizado:", currentMeta); });
            }, 500);
          });
        }
      }
    } catch (error) {
      console.error('Error loading existing metadata:', error);
      this.setState({
        message: `Error cargando metadatos existentes: ${error.message}`,
        messageType: MessageBarType.error,
        showMessage: true
      });
    } finally {
      this.setState({ isLoadingExistingMetadata: false });
    }
  };

  private detectLibraryFromDriveName = (driveId: string): string | null => {
    if (!driveId) return null;
    return null;
  };

  private generateCurrentMetadataFromExisting = (existingMetadata: any, bibliotecaId: string): any => {
    const currentMetadata: any = {};
    switch (bibliotecaId.toUpperCase()) {
      case "DOCUMENTOS_CLIENTES":
        currentMetadata.Cliente = existingMetadata.Cliente || "";
        currentMetadata.Asunto = existingMetadata.Asunto || "";
        currentMetadata.S_Asunto = existingMetadata.S_Asunto || "";
        currentMetadata.Tipo_Doc = existingMetadata.Tipo_Doc || "";
        currentMetadata.S_Tipo = existingMetadata.S_Tipo || "";
        break;

      case "DOCUMENTOS_SOCIOS":
        break;

      case "DOCUMENTOS_ADMIN_RRHH":
        currentMetadata.CarpetaRRHH = existingMetadata.Carpeta1 || existingMetadata.CarpetaRRHH || "";
        break;

      case "DOCUMENTOS_CONSULADO_AUSTRALIA":
        currentMetadata.Nivel1 = existingMetadata.Nivel1 || "";
        currentMetadata.Nivel2 = existingMetadata.Nivel2 || "";
        break;

      case "DOCUMENTOS_CONTADURIA":
        currentMetadata.Tema = existingMetadata.Tema || "";
        currentMetadata.SubTema = existingMetadata.SubTema || "";
        currentMetadata.TipoDoc = existingMetadata.TipoDoc || "";
        break;

      case "DOCUMENTOS_DECLARACIONES_JURADAS":
        currentMetadata.Carpeta1 = existingMetadata.Carpeta1 || "";
        break;

      case "DOCUMENTOS_INTERNO":
        currentMetadata.Carpeta1 = existingMetadata.Carpeta1 || "";
        currentMetadata.Carpeta2 = existingMetadata.Carpeta2 || "";
        currentMetadata.Carpeta3 = existingMetadata.Carpeta3 || "";
        currentMetadata.Carpeta4 = existingMetadata.Carpeta4 || "";
        currentMetadata.Carpeta5 = existingMetadata.Carpeta5 || "";
        currentMetadata.Carpeta6 = existingMetadata.Carpeta6 || "";
        currentMetadata.Carpeta7 = existingMetadata.Carpeta7 || "";
        break;
    }
    const now = new Date().toLocaleDateString("es-ES", {
        day: "numeric",
        month: "numeric",
        year: "numeric",
      });
    currentMetadata.DateCreated = existingMetadata.DateCreated || existingMetadata.Created;
    currentMetadata.DateModified = now;
    currentMetadata.DocID = existingMetadata.DocID;
    currentMetadata.ProcesadoAddIn = existingMetadata.ProcesadoAddIn || "Si";
    currentMetadata.autorEmail = existingMetadata.autorEmail;
    currentMetadata.operadorEmail = existingMetadata.operadorEmail;
    return currentMetadata;
  };

  private updateOfficeDocumentTitle = async (newTitle: string): Promise<void> => {
    try {
      if (this.state.officeApp === "Excel") {
        const canProceed = await this.exitCellEditMode();
        if (!canProceed) return;
      }

      if (this.state.officeApp === "Excel") {
        await Excel.run(async (context) => {
          const workbook = context.workbook;
          const properties = workbook.properties;
          properties.title = newTitle;
          const application = context.workbook.application;
          await context.sync();
          application.calculate(Excel.CalculationType.recalculate);
          await context.sync();
        });

      } else if (this.state.officeApp === "Word") {
        await Word.run(async (context) => {
          const properties = context.document.properties;
          properties.title = newTitle;
          properties.subject = newTitle;
          await context.sync();
        });

      } else if (this.state.officeApp === "PowerPoint") {
        await PowerPoint.run(async (context) => {
          const presentation: any = context.presentation;
          presentation.title = newTitle;
          await context.sync();
        });
      }
    } catch (error) {
      console.warn("No se pudo actualizar el título en Office:", error);
    }
  };

  loadAvailableLibraries = async () => {
    try {
      const siteResponse = await getSiteId(this.props.accessToken, this.SITE_URL);
      if (!siteResponse.success) throw new Error(`Failed to get site: ${siteResponse.error}`);
      const drivesResponse = await getSharePointDrives(this.props.accessToken, siteResponse.data);
      if (!drivesResponse.success) throw new Error(`Failed to get drives: ${drivesResponse.error}`);
      const availableDrives = drivesResponse.data;
      const mappedLibraries = BIBLIOTECAS_DISPONIBLES.filter(lib => {
        // Para DOCUMENTOS_CLIENTES, buscar drives que coincidan con patrones de CLIENTES
        if (lib.id === "DOCUMENTOS_CLIENTES") {
          return availableDrives.some((drive: any) => {
            const driveName = drive.name.toUpperCase();
            return driveName.includes("CLIENTES") || 
                   driveName.includes("DOCUMENTOS CLIENTES") ||
                   driveName.includes("DOCUMENTOS_CLIENTES");
          });
        }
        
        // Para otras bibliotecas, usar la lógica original
        return availableDrives.some((drive: any) =>
          drive.name === lib.id ||
          drive.name.includes(lib.id) ||
          lib.title.includes(drive.name)
        );
      });

      if (mappedLibraries.length > 0) this.setState({ bibliotecas: mappedLibraries });
    } catch (error) {
      console.error('Error loading libraries:', error);
    }
  };

  private attemptSaveDocument = async (): Promise<void> => {
    try {
      if (this.state.officeApp === "Excel") {
        const canProceed = await this.exitCellEditMode();
        if (!canProceed) throw new Error("No se puede proceder mientras Excel esté en modo de edición de celdas");
      }

      if (this.state.officeApp === "Word") {
        await Word.run(async (context) => {
          await context.sync();
        });
      } else if (this.state.officeApp === "Excel") {
        await Excel.run(async (context) => {
          const application = context.workbook.application;
          application.calculate(Excel.CalculationType.full);
          await context.sync();
        });

      } else if (this.state.officeApp === "PowerPoint") {
        await PowerPoint.run(async (context) => { await context.sync(); });
      }
    } catch (error) {
      console.warn("No se pudo sincronizar el documento:", error);
      throw error;
    }
  };

  private getCurrentUserInfo = async (): Promise<{ email: string, name: string, id?: string }> => {
    try {
      const userResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: {
          Authorization: `Bearer ${this.props.accessToken}`,
        },
      });

      if (!userResponse.ok) { throw new Error("No se pudo obtener el usuario actual"); }
      const user = await userResponse.json();

      return {
        email: user.userPrincipalName || user.mail || "",
        name: user.displayName || user.userPrincipalName || ""
      };
    } catch (error) {
      console.error("Error obteniendo información del usuario:", error);
      return {
        email: "unknown@domain.com",
        name: "Unknown User"
      };
    }
  };

  private updateFileMetadata = async (driveId: string, fileId: string, fileName: string, userInfo: { email: string, name: string }, docId?: number, preservedMetadata?: any): Promise<void> => {
    try {
      const now = new Date().toLocaleDateString("es-ES", {
        day: "numeric",
        month: "numeric",
        year: "numeric",
      });
      const metadataToUse = preservedMetadata || this.state.currentMetadata;
      const metadata = {
        Title: fileName,
        ...(this.state.isExistingDocument ? {} : { DateCreated: now }),
        DateModified: now,
        ProcesadoAddIn: "Si",
        autorEmail: userInfo.email,
        operadorEmail: userInfo.email,
        DocID: docId || this.state.currentDocId || null,
        ...metadataToUse
      };
      Object.keys(metadata).forEach(key => {
        if (metadata[key] === undefined || metadata[key] === null || metadata[key] === "") {
          delete metadata[key];
        }
      });

      const updateResponse = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/listItem/fields`,
        {
          method: "PATCH",
          headers: {
            Authorization: `Bearer ${this.props.accessToken}`,
            "Content-Type": "application/json",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify(metadata),
        }
      );


      if (!updateResponse.ok) {
        const errorDetails = await updateResponse.json();
        console.error("Error actualizando metadatos:", errorDetails);
        throw new Error("Error al actualizar los metadatos del archivo");
      }
    } catch (error) {
      console.error("Error en updateFileMetadata:", error);
      throw error;
    }
  };

  getDocumentBlob = async (): Promise<Blob> => {
    const attemptGetFile = (retryCount = 0): Promise<Blob> => {
      return new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(
          Office.FileType.Compressed,
          { sliceSize: 65536 },
          (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              if (retryCount < 2) {
                console.warn(`Intento ${retryCount + 1} falló, reintentando en 1 segundo...`);
                setTimeout(() => attemptGetFile(retryCount + 1).then(resolve).catch(reject), 1000);
                return;
              } else {
                reject(new Error(`No se pudo obtener el archivo después de ${retryCount + 1} intentos: ${result.error?.message || 'Error desconocido'}`));
                return;
              }
            }
            this.processFileSlices(result.value, resolve, reject);
          }
        );
      });
    };
    return attemptGetFile();
  };

  private processFileSlices = (file: Office.File, resolve: (blob: Blob) => void, reject: (error: Error) => void) => {
    const sliceCount = file.sliceCount;
    const slices: Uint8Array[] = [];
    let slicesReceived = 0;
    const getSlice = (index: number) => {
      file.getSliceAsync(index, (sliceResult) => {
        if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
          file.closeAsync();
          return reject(new Error(`Error al obtener slice ${index}: ${sliceResult.error?.message || 'Error desconocido'}`));
        }

        const slice = sliceResult.value;
        slices[index] = new Uint8Array(slice.data as any);
        slicesReceived++;
        if (slicesReceived === sliceCount) {
          const fullData = new Uint8Array(slices.reduce((acc, val) => acc + val.length, 0));
          let offset = 0
          for (const s of slices) {
            fullData.set(s, offset);
            offset += s.length;
          }

          file.closeAsync();
          let mimeType = 'application/octet-stream';
          if (this.state.officeApp === "Excel") {
            mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
          } else if (this.state.officeApp === "Word") {
            mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
          } else if (this.state.officeApp === "PowerPoint") {
            mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
          }
          resolve(new Blob([fullData], { type: mimeType }));
        } else {
          getSlice(index + 1);
        }
      });
    };
    getSlice(0);
  };

  getFileExtension = (): string => {
    switch (this.state.officeApp) {
      case "Excel": return ".xlsx";
      case "Word": return ".docx";
      case "PowerPoint": return ".pptx";
      default: return ".bin";
    }
  };

  private closeOfficeDocument = async (): Promise<boolean> => {
    try {
      if (this.state.officeApp === "Excel") {
        await Excel.run(async (context) => {
          const workbook = context.workbook;
          await context.sync();
          (workbook as any).close();
          await context.sync();
        });
      } else if (this.state.officeApp === "Word") {
        await Word.run(async (context) => {
          const document = context.document;
          await context.sync();
          (document as any).close();
          await context.sync();
        });
      } else if (this.state.officeApp === "PowerPoint") {
        await PowerPoint.run(async (context) => {
          const presentation = context.presentation;
          await context.sync();
          (presentation as any).close();
          await context.sync();
        });
      }
      return true;
    } catch (error) {
      console.error("Error cerrando el documento:", error);
      return false;
    }
  };

  private handleCloseDocumentConfirm = async () => {
    console.log('DEBUG handleCloseDocumentConfirm - Starting');
    console.log('DEBUG handleCloseDocumentConfirm - Current state:', {
      pendingSaveData: this.state.pendingSaveData,
      selectedBiblioteca: this.state.selectedBiblioteca,
      currentFileId: this.state.currentFileId
    });
    this.setState({ showCloseDocumentDialog: false, isSaving: true });

    try {
      console.log('DEBUG handleCloseDocumentConfirm - Calling writeToMetadataQueue');
      await this.writeToMetadataQueue();
      console.log('DEBUG handleCloseDocumentConfirm - writeToMetadataQueue completed successfully');

      // 2. Cerrar el documento
      console.log('DEBUG handleCloseDocumentConfirm - About to close document');
      await this.closeOfficeDocument();
      console.log('DEBUG handleCloseDocumentConfirm - Document closed successfully');

      this.setState({
        isSaving: false,
        pendingSaveData: undefined,
        message: `Solicitud enviada correctamente. Los metadatos se actualizarán automáticamente. El documento se ha cerrado.`,
        messageType: MessageBarType.success,
        showMessage: true
      });
      const closed = await this.closeOfficeDocument();

      if (closed) {
        await new Promise(resolve => setTimeout(resolve, 10000));
        if (this.state.pendingSaveData) {
          await this.updateExistingDocumentMetadata(this.state.pendingSaveData);
        }
      } else {
        throw new Error("No se pudo cerrar el documento");
      }
    } catch (error) {
      console.error("Error en el proceso de cierre y actualización:", error);
      this.setState({
        isSaving: false,
        message: `Error: ${error.message}`,
        messageType: MessageBarType.error,
        showMessage: true,
        pendingSaveData: undefined
      });
    }
  };

  private writeToMetadataQueue = async (): Promise<void> => {
    console.log('DEBUG writeToMetadataQueue - Starting, pendingSaveData:', this.state.pendingSaveData);
    console.log('DEBUG writeToMetadataQueue - Function called successfully');
    try {
      if (!this.state.pendingSaveData) {
        console.log('DEBUG writeToMetadataQueue - No pendingSaveData, throwing error');
        throw new Error("No hay datos para procesar");
      }
      console.log('DEBUG writeToMetadataQueue - pendingSaveData exists, proceeding');

      const siteResponse = await getSiteId(this.props.accessToken, this.SITE_URL);
      if (!siteResponse.success) {
        throw new Error(`Error obteniendo sitio: ${siteResponse.error}`);
      }

      const drivesResponse = await getSharePointDrives(this.props.accessToken, siteResponse.data);
      if (!drivesResponse.success) {
        throw new Error(`Error obteniendo drives: ${drivesResponse.error}`);
      }

      const drives = drivesResponse.data;
      // Use the correct drive name mapping for DOCUMENTOS_CLIENTES
      const driveName = this.getDriveNameFromLibraryId(this.state.selectedBiblioteca);
      const targetDrive = drives.find((drive: any) =>
        drive.name === driveName ||
        drive.id === driveName
      );

      if (!targetDrive) {
        throw new Error("No se encontró la biblioteca");
      }

    // Obtener información del archivo para conseguir el ID único de SharePoint
    let sharepointItemId = null;
    console.log('DEBUG sharepointUniqueId - currentFileId:', this.state.currentFileId);
    console.log('DEBUG sharepointUniqueId - targetDrive.id:', targetDrive.id);
    console.log('DEBUG sharepointUniqueId - targetDrive.name:', targetDrive.name);
    
    if (this.state.currentFileId) {
      try {
        const url = `https://graph.microsoft.com/v1.0/drives/${targetDrive.id}/items/${this.state.currentFileId}/listItem?$select=id`;
        console.log('DEBUG sharepointUniqueId - Fetching URL:', url);
        
        const fileInfoResponse = await fetch(url, {
          headers: { Authorization: `Bearer ${this.props.accessToken}` }
        });
        
        console.log('DEBUG sharepointUniqueId - Response status:', fileInfoResponse.status);
        
        if (fileInfoResponse.ok) {
          const listItem = await fileInfoResponse.json();
          console.log('DEBUG sharepointUniqueId - Response data:', listItem);
          // Este es el ID numérico de la columna ID de SharePoint
          sharepointItemId = listItem.id;
          console.log("SharePoint Item ID obtenido:", sharepointItemId);
        } else {
          const errorText = await fileInfoResponse.text();
          console.error('DEBUG sharepointUniqueId - Error response:', errorText);
        }
      } catch (error) {
        console.warn("No se pudo obtener el ID de SharePoint:", error);
      }
    } else {
      console.log('DEBUG sharepointUniqueId - No currentFileId available');
    }

      // Obtener usuario actual
      const userResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${this.props.accessToken}` }
      });
      const user = await userResponse.json();
      // this.state.currentMetadata.DateModified = now;
      console.log('DEBUG sharepointUniqueId - Final sharepointItemId value:', sharepointItemId);
      const listItem = {
        Title: this.state.currentTitle,
        FileId: this.state.currentFileId,
        DriveId: targetDrive.id,
        Estado: "Pendiente",
        NuevoTitulo: this.state.newTitle.trim(),
        MetadatosJson: JSON.stringify({
          biblioteca: this.state.selectedBiblioteca,
          metadata: this.state.currentMetadata,
          timestamp: new Date().toISOString(),
          sharepointUniqueId: sharepointItemId
        }),
        Usuario: user.userPrincipalName || user.mail || "",
        DocID: this.state.currentDocId ? parseInt(this.state.currentDocId) : null,
        FechaSolicitud: new Date().toISOString()
      };
      console.log('DEBUG sharepointUniqueId - Final listItem to save:', listItem);
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteResponse.data}/lists/${this.METADATA_QUEUE_LIST_ID}/items`,
        {
          method: "POST",
          headers: {
            'Authorization': `Bearer ${this.props.accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ fields: listItem })
        }
      );
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(`Error escribiendo en lista: ${response.status} - ${errorData.error?.message || 'Error desconocido'}`);
      }
      console.log('DEBUG writeToMetadataQueue - Successfully saved to metadata queue');
    } catch (error) {
      console.error("Error en writeToMetadataQueue:", error);
      throw error;
    }
  };

  async addToMetadataQueue(queueItem: {
    title: string;
    fileId: string;
    driveId: string;
    nuevoTitulo?: string;
    metadatosJson: string;
    usuario?: string;
    docId?: number;
    sharepointUniqueId?: string;
  }): Promise<{ success: boolean; error?: string }> {
    try {
      const siteId = await getSiteId(this.props.accessToken, this.SITE_URL);
      const METADATA_QUEUE_LIST_ID = "d81f68e5-fdb6-4555-9739-864878646ae8";

      let sharepointUniqueId = queueItem.sharepointUniqueId;
    if (!sharepointUniqueId && queueItem.fileId) {
      try {
        const fileInfoResponse = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${queueItem.driveId}/items/${queueItem.fileId}?$select=id,sharepointIds`,
          {
            headers: { Authorization: `Bearer ${this.props.accessToken}` }
          }
        );
        
        if (fileInfoResponse.ok) {
          const fileInfo = await fileInfoResponse.json();
          sharepointUniqueId = fileInfo.sharepointIds?.listItemUniqueId || fileInfo.id;
        }
      } catch (error) {
        console.warn("No se pudo obtener el ID único de SharePoint:", error);
        sharepointUniqueId = queueItem.fileId;
      }
    }

      const listItem = {
        Title: queueItem.title,
        FileId: queueItem.fileId,
        DriveId: queueItem.driveId,
        Estado: "Pendiente",
        NuevoTitulo: queueItem.nuevoTitulo || "",
        MetadatosJson: queueItem.metadatosJson,
        Usuario: queueItem.usuario || "",
        DocID: queueItem.docId || null,
        FechaSolicitud: new Date().toISOString(),
        SharePointUniqueId: sharepointUniqueId,
      };

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${METADATA_QUEUE_LIST_ID}/items`,
        {
          method: "POST",
          headers: {
            'Authorization': `Bearer ${this.props.accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            fields: listItem
          })
        }
      );

      if (!response.ok) {
        const errorData = await response.json();
        console.error("Error adding to metadata queue:", errorData);
        return {
          success: false,
          error: `HTTP ${response.status}: ${errorData.error?.message || 'Error desconocido'}`
        };
      }
      return { success: true };
    } catch (error) {
      console.error("Error en addToMetadataQueue:", error);
      return {
        success: false,
        error: error.message || 'Error desconocido'
      };
    }
  }

  private handleCloseDocumentCancel = () => {
    this.setState({
      showCloseDocumentDialog: false,
      pendingSaveData: undefined,
      isSaving: false
    });
  };

  saveToSharePoint = async () => {
    if (!this.state.selectedBiblioteca) {
      this.setState({
        message: "Please select a SharePoint library",
        messageType: MessageBarType.warning,
        showMessage: true
      });
      return;
    }

    if (!this.state.newTitle.trim()) {
      this.setState({
        message: "Please enter a valid title",
        messageType: MessageBarType.warning,
        showMessage: true
      });
      return;
    }

    // Validar que el campo Cliente sea obligatorio para DOCUMENTOS_CLIENTES
    if (this.state.selectedBiblioteca === "DOCUMENTOS_CLIENTES") {
      const metadata = this.generateCurrentMetadataFromExisting(this.state.currentMetadata, this.state.selectedBiblioteca);
      if (!metadata.Cliente || metadata.Cliente === "SIN CLASIFICAR") {
        this.setState({
          message: "El campo Cliente es obligatorio. Por favor seleccione un cliente antes de guardar.",
          messageType: MessageBarType.warning,
          showMessage: true
        });
        return;
      }
    }

    if (this.state.isInCellEditMode || (this.state.officeApp === "Excel")) {
      const canProceed = await this.exitCellEditMode();
      if (!canProceed) {
        return;
      }
    }

    let currentComponentState = null;
    if (this.metadataComponentRef && this.metadataComponentRef.current) {
      const currentInstance = this.metadataComponentRef.current as any;
      currentComponentState = { ...currentInstance.state };
    }

    const preservedState = {
      bibliotecaId: this.state.selectedBiblioteca,
      metadata: { ...this.state.currentMetadata },
      existingMetadata: this.state.existingMetadata ? { ...this.state.existingMetadata } : null,
      componentState: currentComponentState
    };

    try {
      this.setState({ isSaving: true });

      await this.updateTitle();
      await this.attemptSaveDocument();
      const userInfo = await this.getCurrentUserInfo();
      await new Promise(resolve => setTimeout(resolve, 1000));
      const documentBlob = await this.getDocumentBlob();
      const baseFileName = this.state.newTitle.trim();

      console.log('DEBUG createNewDocument - isExistingDocument:', this.state.isExistingDocument, 'currentFileId:', this.state.currentFileId, 'currentDocId:', this.state.currentDocId);
      if ((this.state.isExistingDocument && this.state.currentFileId) || this.state.currentDocId) {
        console.log('DEBUG createNewDocument - Setting pendingSaveData for existing document');
        this.setState({
          showCloseDocumentDialog: true,
          isSaving: false,
          pendingSaveData: {
            documentBlob,
            baseFileName,
            userInfo,
            preservedMetadata: preservedState.metadata
          }
        });
        return;
      } else {
        // Para documentos nuevos, proceder normalmente
        await this.createNewDocument(documentBlob, baseFileName, userInfo, preservedState.metadata);
        this.setState({
          // isSaving: false,
          message: `Documento guardado correctamente en la biblioteca: ${this.state.bibliotecas.find(b => b.id === this.state.selectedBiblioteca)?.title}`,
          messageType: MessageBarType.success,
          showMessage: true
        });
      }

    } catch (error) {
      console.error("Error en saveToSharePoint:", error);
      this.setState({
        isSaving: false,
        message: `Error saving to SharePoint: ${error.toString()}`,
        messageType: MessageBarType.error,
        showMessage: true
      });
    }
  };

  private updateExistingDocumentMetadata = async (saveData: any) => {
    try {
      const { baseFileName, userInfo, preservedMetadata } = saveData;
      let finalDocId = this.state.currentDocId;

      const siteResponse = await getSiteId(this.props.accessToken, this.SITE_URL);
      if (!siteResponse.success) {
        throw new Error(`Failed to get site: ${siteResponse.error}`);
      }

      const drivesResponse = await getSharePointDrives(this.props.accessToken, siteResponse.data);
      if (!drivesResponse.success) {
        throw new Error(`Failed to get drives: ${drivesResponse.error}`);
      }

      const drives = drivesResponse.data;
      const targetDrive = drives.find((drive: any) =>
        drive.name === this.state.selectedBiblioteca ||
        drive.id === this.state.selectedBiblioteca ||
        drive.name.includes(this.state.selectedBiblioteca)
      );

      if (targetDrive && this.state.currentFileId) {
        //await this.updateFileContentDirect(targetDrive.id, this.state.currentFileId, documentBlob);
        const currentBaseName = this.extractBaseTitle(this.state.currentTitle);
        if (currentBaseName !== baseFileName) {
          const newFileName = `${baseFileName}-${finalDocId}${this.getFileExtension()}`;
          await this.renameFileWithDocId(targetDrive.id, this.state.currentFileId, newFileName, parseInt(finalDocId));
        }

        // Actualizar metadatos
        const finalFileName = `${baseFileName}-${finalDocId}${this.getFileExtension()}`;
        await this.updateFileMetadata(targetDrive.id, this.state.currentFileId, finalFileName, userInfo, parseInt(finalDocId), preservedMetadata);

        const newTitleWithDocId = `${baseFileName}-${finalDocId}`;
        this.setState({
          currentTitle: newTitleWithDocId,
          baseTitle: baseFileName,
          originalDocumentSaved: true,
          isSaving: false,
          pendingSaveData: undefined,
          message: `Documento actualizado correctamente. Los metadatos han sido guardados en SharePoint.`,
          messageType: MessageBarType.success,
          showMessage: true
        });
      }
    } catch (error) {
      console.error("Error actualizando documento existente:", error);

      // Detectar si es un error por archivo bloqueado
      if (error.message?.includes("423") || error.message?.includes("resourceLocked")) {
        console.warn("Archivo bloqueado. Guardando cambios para sincronización posterior...");

        this.setState({
          isSaving: false,
          pendingSaveData: undefined,
          pendingSync: true,
          localChanges: {
            currentFileId: this.state.currentFileId,
            currentDocId: this.state.currentDocId,
            newTitle: this.state.newTitle,
            selectedBiblioteca: this.state.selectedBiblioteca,
            currentMetadata: saveData.preservedMetadata,
            isExistingDocument: true,
            componentState: this.state.preservedComponentState
          },
          message: `El archivo está bloqueado. Los cambios se sincronizarán automáticamente cuando reabra el archivo.`,
          messageType: MessageBarType.warning,
          showMessage: true
        });

        return;
      }

      this.setState({
        isSaving: false,
        pendingSaveData: undefined,
        message: `Error actualizando documento: ${error.message}`,
        messageType: MessageBarType.error,
        showMessage: true
      });
    }
  };

  private attemptBackgroundSync = async (): Promise<void> => {
    if (!this.state.pendingSync || !this.state.localChanges) {
      return;
    }

    const localChanges = this.state.localChanges;
    const userInfo = await this.getCurrentUserInfo();

    if (localChanges.isExistingDocument && localChanges.currentFileId) {
      const siteResponse = await getSiteId(this.props.accessToken, this.SITE_URL);
      if (!siteResponse.success) throw new Error("Site no disponible");

      const drivesResponse = await getSharePointDrives(this.props.accessToken, siteResponse.data);
      if (!drivesResponse.success) throw new Error("Drives no disponibles");

      const drives = drivesResponse.data;
      const targetDrive = drives.find((drive: any) =>
        drive.name === localChanges.selectedBiblioteca ||
        drive.id === localChanges.selectedBiblioteca ||
        drive.name.includes(localChanges.selectedBiblioteca)
      );

      if (!targetDrive) throw new Error("Drive no encontrado");
      const checkResponse = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${targetDrive.id}/items/${localChanges.currentFileId}`,
        {
          headers: {
            Authorization: `Bearer ${this.props.accessToken}`,
          },
        }
      );

      if (!checkResponse.ok) {
        throw new Error("Archivo aún no disponible");
      }

      const finalFileName = `${localChanges.newTitle}-${localChanges.currentDocId}${this.getFileExtension()}`;
      await this.updateFileMetadata(
        targetDrive.id,
        localChanges.currentFileId,
        finalFileName,
        userInfo,
        parseInt(localChanges.currentDocId || "0"),
        localChanges.currentMetadata
      );

      this.setState({
        pendingSync: false,
        localChanges: {},
        syncRetryCount: 0,
        message: "Cambios sincronizados exitosamente con SharePoint.",
        messageType: MessageBarType.success,
        showMessage: true
      });

      if (this.metadataComponentRef && this.metadataComponentRef.current && localChanges.componentState) {
        setTimeout(() => {
          const currentInstance = this.metadataComponentRef.current as any;
          if (currentInstance.restoreState) {
            currentInstance.restoreState(localChanges.componentState);
          }
        }, 100);
      }
    }
  };

  public forceSyncNow = async (): Promise<void> => {
    if (!this.state.pendingSync) {
      this.setState({
        message: "No hay cambios pendientes para sincronizar.",
        messageType: MessageBarType.info,
        showMessage: true
      });
      return;
    }
    this.setState({ isSaving: true });
    try {
      await this.attemptBackgroundSync();
    } catch (error) {
      this.setState({
        isSaving: false,
        message: "No se pudo sincronizar en este momento. El archivo puede estar aún en uso: " + error.message,
        messageType: MessageBarType.warning,
        showMessage: true
      });
    }
  };

  private saveToSharePointWithRetry = async (retryCount: number = 0, preservedState?: any): Promise<void> => {
    try {
      this.setState({ isSaving: true });
      const currentBiblioteca = preservedState?.bibliotecaId || this.state.selectedBiblioteca;
      const currentMetadata = preservedState?.metadata || { ...this.state.currentMetadata };
      const currentExistingMetadata = preservedState?.existingMetadata || this.state.existingMetadata;
      const componentState = preservedState?.componentState;
      await this.updateTitle();
      await this.attemptSaveDocument();
      const userInfo = await this.getCurrentUserInfo();
      await new Promise(resolve => setTimeout(resolve, 3000));
      const documentBlob = await this.getDocumentBlob();
      const baseFileName = this.state.newTitle.trim();
      let finalDocId = this.state.currentDocId;

      if ((this.state.isExistingDocument && this.state.currentFileId) || this.state.currentDocId) {
        await this.updateExistingDocument(documentBlob, baseFileName, userInfo, finalDocId, currentMetadata);
      } else {
        await this.createNewDocument(documentBlob, baseFileName, userInfo, currentMetadata);
        return;
      }

      this.setState({
        // isSaving: false,
        message: `Documento ${this.state.isExistingDocument ? 'actualizado' : 'guardado'} correctamente en la biblioteca: ${this.state.bibliotecas.find(b => b.id === currentBiblioteca)?.title}. El documento ya está guardado en SharePoint.`,
        messageType: MessageBarType.success,
        showMessage: true,
        selectedBiblioteca: currentBiblioteca,
        currentMetadata: currentMetadata,
        existingMetadata: currentExistingMetadata,
        preservedComponentState: componentState
      }, () => {
        setTimeout(() => {
          if (this.metadataComponentRef && this.metadataComponentRef.current && componentState) {
            const currentInstance = this.metadataComponentRef.current as any;
            if (currentInstance.restoreState) {
              currentInstance.restoreState(componentState);
            }
          }
        }, 100);
      });
    } catch (error) {
      if (preservedState) {
        this.setState({
          selectedBiblioteca: preservedState.bibliotecaId,
          currentMetadata: preservedState.metadata,
          existingMetadata: preservedState.existingMetadata,
          preservedComponentState: preservedState.componentState
        });
      }

      if (error.message.includes("423") || error.message.includes("resourceLocked")) {
        if (retryCount < this.MAX_RETRY_ATTEMPTS) {
          this.setState({
            // isSaving: false,
            message: `El archivo está siendo utilizado por otro proceso. Reintentando en ${(retryCount + 1) * 2} segundos... (${retryCount + 1}/${this.MAX_RETRY_ATTEMPTS})`,
            messageType: MessageBarType.warning,
            showMessage: true
          });
          setTimeout(() => {
            this.saveToSharePointWithRetry(retryCount + 1, preservedState);
          }, (retryCount + 1) * 2000);
          return;
        } else {
          this.setState({
            isSaving: false,
            message: "Error: El archivo está siendo utilizado por otro proceso y no se pudo guardar después de varios intentos. Por favor, cierre otras instancias del archivo e intente nuevamente.",
            messageType: MessageBarType.error,
            showMessage: true
          });
          return;
        }
      }
      let errorMessage = "Error saving to SharePoint: ";
      if (error.message.includes("No se pudo obtener el archivo")) {
        errorMessage += "El documento puede estar siendo editado. Guarda los cambios manualmente primero (Ctrl+S) y luego intenta nuevamente.";
      } else if (error.message.includes("InvalidOperationInCellEditMode")) {
        errorMessage += "Excel está en modo de edición de celdas. Presione Enter, Tab o Escape para salir del modo de edición.";
      } else { errorMessage += error.toString(); }
      this.setState({
        isSaving: false,
        message: errorMessage,
        messageType: MessageBarType.error,
        showMessage: true
      });
    }
  };

  private updateExistingDocument = async (_documentBlob: Blob, baseFileName: string, userInfo: any, finalDocId: string, preservedMetadata?: any) => {
    const siteResponse = await getSiteId(this.props.accessToken, this.SITE_URL);
    if (siteResponse.success) {
      const drivesResponse = await getSharePointDrives(this.props.accessToken, siteResponse.data);
      if (drivesResponse.success) {
        const drives = drivesResponse.data;
        const targetDrive = drives.find((drive: any) =>
          drive.name === this.state.selectedBiblioteca ||
          drive.id === this.state.selectedBiblioteca ||
          drive.name.includes(this.state.selectedBiblioteca)
        );

        if (targetDrive && this.state.currentFileId) {
          //await this.updateFileContentWithRetry(targetDrive.id, this.state.currentFileId, documentBlob);
          const currentBaseName = this.extractBaseTitle(this.state.currentTitle);
          if (currentBaseName !== baseFileName) {
            const newFileName = `${baseFileName}-${finalDocId}${this.getFileExtension()}`;
            await this.renameFileWithDocId(targetDrive.id, this.state.currentFileId, newFileName, parseInt(finalDocId));
          }
          const finalFileName = `${baseFileName}-${finalDocId}${this.getFileExtension()}`;
          const metadataToUse = preservedMetadata || this.state.currentMetadata;
          await this.updateFileMetadata(targetDrive.id, this.state.currentFileId, finalFileName, userInfo, parseInt(finalDocId), metadataToUse);
          const newTitleWithDocId = `${baseFileName}-${finalDocId}`;
          this.setState({
            currentTitle: newTitleWithDocId,
            baseTitle: baseFileName,
            originalDocumentSaved: true
          });
          await this.updateOfficeDocumentTitle(newTitleWithDocId);
        }
      }
    }
  };

  private getDriveNameFromLibraryId = (libraryId: string): string => {
    // Mapear IDs internos a nombres reales de drives
    const driveMapping: { [key: string]: string } = {
      "DOCUMENTOS_CLIENTES": "DOCUMENTOS_CLIENTES",
      "DOCUMENTOS_CLIENTES_V2": "DOCUMENTOS_CLIENTES", // Mapear V2 a la versión real
      "DOCUMENTOS_SOCIOS": "DOCUMENTOS_SOCIOS",
      "DOCUMENTOS_ADMIN_RRHH": "DOCUMENTOS_ADMIN_RRHH",
      "DOCUMENTOS_CONSULADO_AUSTRALIA": "DOCUMENTOS_CONSULADO_AUSTRALIA",
      "DOCUMENTOS_CONTADURIA": "DOCUMENTOS_CONTADURIA",
      "DOCUMENTOS_DECLARACIONES_JURADAS": "DOCUMENTOS_DECLARACIONES_JURADAS",
      "DOCUMENTOS_INTERNO": "DOCUMENTOS_INTERNO"
    };
    
    return driveMapping[libraryId] || libraryId;
  };

  private createNewDocument = async (documentBlob: Blob, baseFileName: string, userInfo: any, preservedMetadata?: any) => {
    const fileName = `${baseFileName}${this.getFileExtension()}`;
    
    // Mapear el ID interno de la biblioteca al nombre real del drive
    const driveName = this.getDriveNameFromLibraryId(this.state.selectedBiblioteca);
    
    
    const result: SharePointResponse = await uploadToSharePoint(
      this.props.accessToken,
      this.SITE_URL,
      driveName,
      fileName,
      documentBlob
    );

    if (result.success && result.data) {
      const siteResponse = await getSiteId(this.props.accessToken, this.SITE_URL);
      if (siteResponse.success) {
        const drivesResponse = await getSharePointDrives(this.props.accessToken, siteResponse.data);
        if (drivesResponse.success) {
          const drives = drivesResponse.data;
          const targetDrive = drives.find((drive: any) => {
            // Para DOCUMENTOS_CLIENTES y DOCUMENTOS_CLIENTES_V2, buscar exactamente DOCUMENTOS_CLIENTES
            if (this.state.selectedBiblioteca === "DOCUMENTOS_CLIENTES" || this.state.selectedBiblioteca === "DOCUMENTOS_CLIENTES_V2") {
              return drive.name === "DOCUMENTOS_CLIENTES";
            }
            
            // Para otros drives, usar la lógica original
            return drive.name === this.state.selectedBiblioteca ||
                   drive.id === this.state.selectedBiblioteca ||
                   drive.name.includes(this.state.selectedBiblioteca);
          });

          if (targetDrive && result.data.id) {
            const listItemResponse = await fetch(
              `https://graph.microsoft.com/v1.0/drives/${targetDrive.id}/items/${result.data.id}/listItem?$select=Id`,
              {
                headers: {
                  Authorization: `Bearer ${this.props.accessToken}`,
                },
              }
            );

            if (listItemResponse.ok) {
              const listItem = await listItemResponse.json();
              const docId = listItem.id;
              const finalFileName = await this.renameFileWithDocId(targetDrive.id, result.data.id, fileName, docId);
              const metadataToUse = preservedMetadata || this.state.currentMetadata;
              await this.updateFileMetadata(targetDrive.id, result.data.id, finalFileName, userInfo, docId, metadataToUse);
              const newTitleWithDocId = `${baseFileName}-${docId}`;
              this.setState({
                currentTitle: newTitleWithDocId,
                baseTitle: baseFileName,
                currentDocId: docId.toString(),
                originalDocumentSaved: true
              });
              await this.updateOfficeDocumentTitle(newTitleWithDocId);
              await this.markDocumentAsSaved();
            } else { throw new Error("No se pudo obtener el DocID del archivo"); }
          }
        }
      }
      this.setState({
        // isSaving: false,
        message: `Documento guardado correctamente en la biblioteca: ${this.state.bibliotecas.find(b => b.id === this.state.selectedBiblioteca)?.title}. El documento ya está guardado en SharePoint.`,
        messageType: MessageBarType.success,
        showMessage: true,
        isExistingDocument: true,
        currentFileId: result.data.id
      });
    } else { throw new Error(result.error || "Unknown error occurred"); }
  };

  private updateFileContentWithRetry = async (driveId: string, fileId: string, documentBlob: Blob, retryCount: number = 0): Promise<void> => {
    try {
      const updateResponse = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`,
        {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${this.props.accessToken}`,
            "Content-Type": "application/octet-stream",
          },
          body: documentBlob,
        }
      );

      if (!updateResponse.ok) {
        if (updateResponse.status === 423 && retryCount < this.MAX_RETRY_ATTEMPTS) {
          await new Promise(resolve => setTimeout(resolve, (retryCount + 1) * 2000));
          return this.updateFileContentWithRetry(driveId, fileId, documentBlob, retryCount + 1);
        }
        throw new Error(`Error al actualizar el contenido del archivo: ${updateResponse.status} - ${await updateResponse.text()}`);
      }
    } catch (error) {
      console.error("Error en updateFileContentWithRetry:", error);
      throw error;
    }
  };

  private markDocumentAsSaved = async (): Promise<void> => {
    try {
      if (this.state.officeApp === "Excel") {
        await Excel.run(async (context: any) => {
          const workbook = context.workbook;
          try {
            await context.document.save();
          } catch (saveError) {
            const application = context.workbook.application;
            const sheet = workbook.worksheets.getActiveWorksheet();
            const cell = sheet.getRange("A1");
            cell.load("value");
            await context.sync();
            const originalValue = cell.value;
            cell.value = [["TEMP"]];
            await context.sync();
            cell.value = originalValue;
            await context.sync();
            application.calculate(Excel.CalculationType.full);
            await context.sync();
          }
        });
      } else if (this.state.officeApp === "Word") {
        await Word.run(async (context) => {
          try {
            await context.document.save();
          } catch (saveError) {
            const document = context.document;
            const body = document.body;
            body.insertText(" ", Word.InsertLocation.end);
            await context.sync();
            const range = body.getRange();
            range.load("text");
            await context.sync();
            const text = range.text;
            if (text.endsWith(" ")) {
              range.insertText(text.substring(0, text.length - 1), Word.InsertLocation.replace);
              await context.sync();
            }
          }
        });
      } else if (this.state.officeApp === "PowerPoint") {
        await PowerPoint.run(async (context) => {
          try {
            await context.sync();
          } catch (saveError) {
            console.log("PowerPoint: usando método de sincronización");
          }
        });
      }
      this.setState({ originalDocumentSaved: true });
      await new Promise(resolve => setTimeout(resolve, 1000));
    } catch (error) {
      console.warn("No se pudo marcar el documento como guardado:", error);
    }
  };

  private renameFileWithDocId = async (driveId: string, fileId: string, originalFileName: string, docId: number): Promise<string> => {
    try {
      const fileExtension = this.getFileExtension();
      const baseFileName = originalFileName.replace(fileExtension, '');
      const docIdPattern = /-\d+$/;
      const cleanBaseFileName = baseFileName.replace(docIdPattern, '');
      const newFileName = `${cleanBaseFileName}-${docId}${fileExtension}`;
      const renameResponse = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}`,
        {
          method: "PATCH",
          headers: {
            Authorization: `Bearer ${this.props.accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            name: newFileName
          }),
        }
      );
      if (!renameResponse.ok) {
        const errorDetails = await renameResponse.json();
        console.error("Error renombrando archivo:", errorDetails);
        throw new Error("Error al renombrar el archivo con DocID");
      }
      return newFileName;
    } catch (error) {
      console.error("Error en renameFileWithDocId:", error);
      throw error;
    }
  };

  dismissMessage = () => {
    this.setState({ showMessage: false });
  };

  // private handleForceSyncClick = () => {
  //   this.handleForceSync(true);
  // };

  private handleForceSync = async (showMessage: boolean = true) => {
    this.setState({ isSyncing: true });
    try {
      // Limpiar cache del servicio de datos
      if (this.dataService && this.dataService.clearCache) {
        this.dataService.clearCache();
      }
      
      // Forzar actualización del cache
      if (this.dataService && this.dataService.forceRefreshCache) {
        await this.dataService.forceRefreshCache();
      }
      
      // Recargar datos del componente de metadatos
      if (this.metadataComponentRef && this.metadataComponentRef.current) {
        const currentInstance = this.metadataComponentRef.current as any;
        if (currentInstance.refreshData) {
          await currentInstance.refreshData();
        } else if (currentInstance.loadDataForLibrary) {
          await currentInstance.loadDataForLibrary();
        }
      }
      
      // Recargar bibliotecas disponibles
      await this.loadAvailableLibraries();
      
      if (showMessage) {
        this.setState({
          message: "Datos sincronizados correctamente. Los metadatos se han actualizado.",
          messageType: MessageBarType.success,
          showMessage: true
        });
      }
    } catch (error) {
      console.error("Error al sincronizar datos:", error);
      if (showMessage) {
        this.setState({
          message: "Error al sincronizar datos. Intente nuevamente.",
          messageType: MessageBarType.error,
          showMessage: true
        });
      }
    } finally {
      this.setState({ isSyncing: false });
    }
  };

  render() {
    const {
      isLoading,
      isSaving,
      newTitle,
      message,
      messageType,
      showMessage,
      selectedBiblioteca,
      bibliotecas,
      existingMetadata,
      isLoadingExistingMetadata,
      isInCellEditMode,
      pendingSync,
      showCloseDocumentDialog,
      // isSyncing
    } = this.state;

    const bibliotecaOptions: IDropdownOption[] = bibliotecas.map(biblioteca => ({
      key: biblioteca.id,
      text: biblioteca.title
    }));

    return (
      <div className="ms-welcome__main">
        {showMessage && (
          <MessageBar
            messageBarType={messageType}
            onDismiss={this.dismissMessage}
            dismissButtonAriaLabel="Close"
          >
            {message}
          </MessageBar>
        )}
        {/* {selectedBiblioteca === "DOCUMENTOS_CLIENTES" && this.renderCacheControls()} */}
        {selectedBiblioteca === "DOCUMENTOS_CLIENTES"}
        {/* NUEVO: Indicador de sincronización pendiente */}
        {pendingSync && (
          <MessageBar
            messageBarType={MessageBarType.info}
            isMultiline={true}
            actions={
              <div>
                <PrimaryButton
                  onClick={this.forceSyncNow}
                  disabled={isSaving}
                  text="Intentar sincronizar ahora"
                  style={{ marginRight: '10px' }}
                />
              </div>
            }
          >
            <strong>🔄 Sincronización pendiente:</strong> Los cambios están guardados localmente y se sincronizarán automáticamente cuando el archivo esté disponible.
          </MessageBar>
        )}

        {isInCellEditMode && (
          <MessageBar
            messageBarType={MessageBarType.severeWarning}
            isMultiline={true}
          >
            <strong>Excel en modo de edición:</strong> Presione <strong>Enter</strong>, <strong>Tab</strong> o <strong>Escape</strong> para salir del modo de edición de celdas, luego haga clic en "Refrescar".
          </MessageBar>
        )}

        {/* NUEVO: Diálogo de confirmación para cerrar documento */}
        <Dialog
          hidden={!showCloseDocumentDialog}
          onDismiss={this.handleCloseDocumentCancel}
          dialogContentProps={{
            type: DialogType.normal,
            title: '📄 Actualizar Metadatos',
            subText: 'Para actualizar los metadatos de este documento existente, es necesario cerrarlo temporalmente. El documento se cerrará, se actualizarán los metadatos en SharePoint, y podrá volver a abrirlo desde SharePoint. ¿Desea continuar?'
          }}
          modalProps={{
            isBlocking: true,
            isDarkOverlay: true
          }}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={this.handleCloseDocumentConfirm}
              text="Sí, cerrar y actualizar"
              disabled={isSaving}
            />
            <DefaultButton
              onClick={this.handleCloseDocumentCancel}
              text="Cancelar"
              disabled={isSaving}
            />
          </DialogFooter>
        </Dialog>

        {isLoading ? (
          <div style={{ textAlign: 'center', padding: '20px' }}>
            <Spinner size={SpinnerSize.large} label="Loading document information..." />
          </div>
        ) : (
          <div>
            <TextField
              label="Título del Documento"
              value={newTitle}
              onChange={this.handleTitleChange}
              placeholder="Enter document title..."
              multiline={false}
              autoComplete="off"
              className="ms-u-marginBottom-16"
              required
              disabled={isInCellEditMode}
              description="Ingrese solo el título base, sin extensión ni DocID."
            />

            <Dropdown
              label="Biblioteca"
              placeholder="Seleccione una biblioteca..."
              options={bibliotecaOptions}
              selectedKey={selectedBiblioteca}
              onChange={this.handleBibliotecaChange}
              className="ms-u-marginBottom-16"
              required
              disabled={isInCellEditMode}
            />

            {isLoadingExistingMetadata && (
              <div style={{ textAlign: 'center', padding: '10px' }}>
                <Spinner size={SpinnerSize.small} label="Cargando metadatos existentes..." />
              </div>
            )}

            {selectedBiblioteca && !isInCellEditMode && (
              <MetadataComponents
                ref={this.metadataComponentRef}
                bibliotecaId={selectedBiblioteca}
                dataService={this.dataService}
                onMetadataChange={this.handleMetadataChange}
                isLoading={false}
                existingMetadata={existingMetadata}
                preservedState={this.state.preservedComponentState}
                key={`metadata-${selectedBiblioteca}`}
              />
            )}

            <div className="ms-u-marginBottom-24">
              <PrimaryButton
                text={isSaving ? "Guardar en SharePoint" : pendingSync ? "Actualizar cambios locales" : "Guardar en SharePoint"}
                onClick={this.saveToSharePoint}
                disabled={isLoading || isSaving || !newTitle.trim() || !selectedBiblioteca || isInCellEditMode}
                style={{
                  marginRight: '10px',
                  backgroundColor: pendingSync ? '#ff8c00' : '#0078d4',
                  marginTop: '20px',
                  marginBottom: '20px'
                }}
                iconProps={isSaving ? { iconName: 'Sync' } : pendingSync ? { iconName: 'CloudUpload' } : { iconName: 'CloudUpload' }}
              />

              {/* NUEVO: Botón para forzar sincronización */}
              {pendingSync && (
                <PrimaryButton
                  text={isSaving ? "Sincronizando..." : "Sincronizar ahora"}
                  onClick={this.forceSyncNow}
                  disabled={isLoading || isSaving}
                  style={{
                    marginRight: '10px',
                    backgroundColor: '#107c10',
                    marginTop: '20px',
                    marginBottom: '20px'
                  }}
                  iconProps={isSaving ? { iconName: 'Sync' } : { iconName: 'CloudUpload' }}
                />
              )}
            </div>

            {/* Botón de sincronización de datos 
            <div style={{ textAlign: 'center', marginBottom: '16px' }}>
              <DefaultButton
                text={isSyncing ? "Sincronizando..." : "🔄 Sincronizar datos"}
                onClick={this.handleForceSyncClick}
                disabled={isSyncing || isLoading || isInCellEditMode}
                iconProps={isSyncing ? { iconName: 'Sync' } : { iconName: 'Sync' }}
                style={{
                  fontSize: '12px',
                  padding: '6px 16px',
                  minWidth: 'auto',
                  backgroundColor: '#f3f2f1',
                  borderColor: '#d2d0ce'
                }}
              />
            </div>*/}
          </div>
        )}
      </div>
    );
  }
}