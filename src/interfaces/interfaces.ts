import { SharePointDataService } from "../services/sharepoint-data.service";
import { MessageBarType } from "office-ui-fabric-react";

export interface Biblioteca {
  id: string;
  title: string;
  listId?: string;
}

export interface CommonProps {
  id: string;
  title: string;
}

export interface MetadataState {
  // CLIENTES
  clientes: CommonProps[];
  clientesFiltered: CommonProps[];
  selectedCliente: CommonProps | null;
  clienteSearchText: string;
  isLoadingClientes: boolean;
  isLoadingAsuntos: boolean;
  isLoadingSubasuntos: boolean;
  isLoadingSubtipos: boolean;

  asuntos: CommonProps[];
  selectedAsunto: CommonProps | null;
  subasuntos: CommonProps[];
  selectedSubasunto: CommonProps | null;
  tiposDocumento: CommonProps[];
  selectedTipoDocumento: CommonProps | null;
  subTiposDocumento: CommonProps[];
  selectedSubTipoDocumento: CommonProps | null;

  // ADMINISTRACION RRHH
  carpetaRRHHOptions: CommonProps[];
  selectedCarpetaRRHH: CommonProps | null;

  // CONSULADO AUSTRALIA
  nivel1Options: CommonProps[];
  selectedNivel1: CommonProps | null;
  nivel2Options: CommonProps[];
  selectedNivel2: CommonProps | null;

  // CONTADURIA
  temasContaduria: CommonProps[];
  selectedTema: CommonProps | null;
  subtemasContaduria: CommonProps[];
  selectedSubtema: CommonProps | null;
  tiposDocumentoContaduria: CommonProps[];
  selectedTipoDocumentoContaduria: CommonProps | null;

  // DJ PROFESIONALES
  carpetaDJOptions: CommonProps[];
  selectedCarpetaDJ: CommonProps | null;

  // INTERNO
  carpeta1Options: CommonProps[];
  selectedCarpeta1: CommonProps | null;
  carpeta2Options: CommonProps[];
  selectedCarpeta2: CommonProps | null;
  carpeta3Options: CommonProps[];
  selectedCarpeta3: CommonProps | null;
  carpeta4Options: CommonProps[];
  selectedCarpeta4: CommonProps | null;
  carpeta5Options: CommonProps[];
  selectedCarpeta5: CommonProps | null;
  carpeta6Options: CommonProps[];
  selectedCarpeta6: CommonProps | null;
  carpeta7Options: CommonProps[];
  selectedCarpeta7: CommonProps | null;

  isLoadingData: boolean;
  hasLoadedExistingMetadata: boolean;
}

export interface SharePointResponse {
  success: boolean;
  data?: any;
  error?: string;
}

export interface DriveInfo {
  id: string;
  name: string;
  driveType: string;
}

export interface MetadataComponentsProps {
  bibliotecaId: string;
  dataService: SharePointDataService;
  onMetadataChange: (metadata: any) => void;
  isLoading?: boolean;
  existingMetadata?: any;
  preservedState?: any;
}

export interface DocumentTitleEditorState {
  currentTitle: string;
  newTitle: string;
  baseTitle: string;
  isLoading: boolean;
  isSaving: boolean;
  message: string;
  messageType: MessageBarType;
  showMessage: boolean;
  officeApp: string;
  selectedBiblioteca: string;
  bibliotecas: Biblioteca[];
  currentMetadata: any;
  existingMetadata: any;
  isExistingDocument: boolean;
  currentFileId: string;
  currentDocId: string;
  isLoadingExistingMetadata: boolean;
  originalDocumentSaved: boolean;
  isInCellEditMode: boolean;
  preservedComponentState: any;

  pendingSync?: boolean;
  localChanges?: any;
  lastSyncAttempt?: Date | null;
  syncRetryCount?: number;
  showCloseDocumentDialog: boolean;
  pendingSaveData?: {
    documentBlob: Blob;
    baseFileName: string;
    userInfo: any;
    preservedMetadata?: any;
  };
  isSyncing: boolean;
}

export interface DocumentTitleEditorProps {
  logout: () => void;
  accessToken: string;
}

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