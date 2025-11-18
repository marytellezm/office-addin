import { Biblioteca } from "../interfaces/interfaces";

// sharepoint-config.ts
export const SHAREPOINT_CONFIG = {
  siteUrl: "https://hughesandhughesuy.sharepoint.com/sites/GestorDocumental",
  lists: {
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
  },
  pageSize: 200
};

export const BIBLIOTECAS_DISPONIBLES: Biblioteca[] = [
  { id: "DOCUMENTOS_CLIENTES", title: "CLIENTES", listId: "d27e0216-23fb-405a-a567-677561e21701" },
  { id: "DOCUMENTOS_SOCIOS", title: "SOCIOS" },
  { id: "DOCUMENTOS_ADMIN_RRHH", title: "ADMINISTRACION RRHH" },
  { id: "DOCUMENTOS_CONSULADO_AUSTRALIA", title: "CONSULADO DE AUSTRALIA" },
  { id: "DOCUMENTOS_CONTADURIA", title: "CONTADURIA" },
  { id: "DOCUMENTOS_DECLARACIONES_JURADAS", title: "DJ PROFESIONALES" },
  { id: "DOCUMENTOS_INTERNO", title: "INTERNO", listId: "cffdd944-add5-4314-bc9c-a40e5c1785f1" }
];

export const METADATA_FIELDS = {
  DOCUMENTOS_CLIENTES: ["Cliente", "Asunto", "S_Asunto", "Tipo_Doc", "S_Tipo"],
  DOCUMENTOS_SOCIOS: ["Contratos"],
  DOCUMENTOS_ADMIN_RRHH: ["Carpeta1"],
  DOCUMENTOS_CONSULADO_AUSTRALIA: ["Nivel1", "Nivel2"],
  DOCUMENTOS_CONTADURIA: ["Tema", "SubTema", "TipoDoc"],
  DOCUMENTOS_DECLARACIONES_JURADAS: ["Carpeta1"],
  DOCUMENTOS_INTERNO: ["Carpeta1", "Carpeta2", "Carpeta3", "Carpeta4", "Carpeta5", "Carpeta6", "Carpeta7"],
};