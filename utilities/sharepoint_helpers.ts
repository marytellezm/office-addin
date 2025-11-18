/*
 * SharePoint integration utilities for Office Add-in - VERSIÓN CON NOMBRE DERIVADO
 */
import { DriveInfo, SharePointResponse } from "../src/interfaces/interfaces";

export const getGraphData = async (url: string, accessToken: string): Promise<any> => {
  const response = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json'
    }
  });

  if (!response.ok) {
    throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
  }

  return await response.json();
};

export const getSiteId = async (accessToken: string, siteUrl: string): Promise<SharePointResponse> => {
  try {
    let graphUrl: string;

    if (siteUrl.includes('/sites/')) {
      const hostname = 'hughesandhughesuy.sharepoint.com';
      const sitePath = '/sites/GestorDocumental';
      graphUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`;
    } else {
      graphUrl = siteUrl;
    }

    const data = await getGraphData(graphUrl, accessToken);
    return { success: true, data: data.id };
  } catch (error) {
    console.error('getSiteId error:', error);
    return { success: false, error: error.toString() };
  }
};

export const getSharePointDrives = async (
  accessToken: string,
  siteId: string
): Promise<SharePointResponse> => {
  try {
    const data = await getGraphData(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      accessToken
    );

    return { success: true, data: data.value };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
};

export const uploadToSharePointByDriveId = async (
  accessToken: string,
  driveId: string,
  fileName: string,
  fileBlob: Blob
): Promise<SharePointResponse> => {
  try {
    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(fileName)}:/content`;

    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/octet-stream'
      },
      body: fileBlob
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Upload failed: ${response.status} - ${errorText}`);
    }

    const result = await response.json();

    return { success: true, data: result };
  } catch (error) {
    console.error('SharePoint upload error:', error);
    return { success: false, error: error.toString() };
  }
};

export const uploadToSharePoint = async (
  accessToken: string,
  siteUrl: string,
  libraryId: string,
  fileName: string,
  fileBlob: Blob
): Promise<SharePointResponse> => {
  try {
    const siteResponse = await getSiteId(accessToken, siteUrl);
    if (!siteResponse.success) {
      throw new Error(`Failed to get site ID: ${siteResponse.error}`);
    }

    const siteId = siteResponse.data;
    const drivesResponse = await getSharePointDrives(accessToken, siteId);
    if (!drivesResponse.success) {
      throw new Error(`Failed to get drives: ${drivesResponse.error}`);
    }

    const drives = drivesResponse.data;
    const targetDrive = drives.find((drive: DriveInfo) => {
      // Para DOCUMENTOS_CLIENTES y DOCUMENTOS_CLIENTES_V2, buscar exactamente DOCUMENTOS_CLIENTES
      if (libraryId === "DOCUMENTOS_CLIENTES" || libraryId === "DOCUMENTOS_CLIENTES_V2") {
        return drive.name === "DOCUMENTOS_CLIENTES";
      }
      
      // Para otros drives, usar la lógica original
      return drive.name === libraryId ||
             drive.id === libraryId ||
             drive.name.includes(libraryId);
    });

    if (!targetDrive) {
      throw new Error(`Library '${libraryId}' not found. Available drives: ${drives.map((d: DriveInfo) => d.name).join(', ')}`);
    }

    return await uploadToSharePointByDriveId(accessToken, targetDrive.id, fileName, fileBlob);
  } catch (error) {
    console.error('SharePoint upload error:', error);
    return { success: false, error: error.toString() };
  }
};

export const findDocumentLibrary = async (
  accessToken: string,
  siteUrl: string,
  fileName: string,
  onLoadingChange?: (isLoading: boolean) => void,
  // preferredLibrary?: string,
  fileId?: string
): Promise<SharePointResponse> => {
  const MAX_RETRIES = 3;
  const RETRY_DELAY = 1000; // 1 second

  const attemptSearch = async (retryCount: number = 0): Promise<SharePointResponse> => {
    try {
      if (onLoadingChange) onLoadingChange(true);

      // Get site ID
      const siteResponse = await getSiteId(accessToken, siteUrl);
      if (!siteResponse.success) {
        throw new Error(`Failed to get site: ${siteResponse.error}`);
      }
      const siteId = siteResponse.data;

      // Get all drives
      const drivesResponse = await getSharePointDrives(accessToken, siteId);
      if (!drivesResponse.success) {
        throw new Error(`Failed to get drives: ${drivesResponse.error}`);
      }
      const drives = drivesResponse.data;

      if (!siteUrl || siteUrl.includes('file://') || !siteUrl.includes('hughesandhughesuy.sharepoint.com')) {
        return {
          success: false,
          error: `Document "${fileName}" is local or new`,
          data: { isNewDocument: true }
        };
      }

      // Fetch the full document URL if siteUrl is the root or invalid
      let effectiveUrl = siteUrl;
      let urlDerivedFileName: string | null = null;
      if (siteUrl === "https://hughesandhughesuy.sharepoint.com/sites/GestorDocumental" || !siteUrl.includes('/Documentos')) {
        const result: any = await new Promise((resolve) => {
          Office.context.document.getFilePropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value.url) {
              effectiveUrl = result.value.url;
              const url = new URL(effectiveUrl);
              const pathSegments = url.pathname.split('/').filter(segment => segment);
              const siteIndex = pathSegments.indexOf('sites');
              if (siteIndex !== -1 && pathSegments.length > siteIndex + 2) {
                urlDerivedFileName = decodeURIComponent(pathSegments[pathSegments.length - 1]);
              }
            } else {
              console.log("No se pudo obtener URL completa:", result.error ? result.error.message : "Unknown error");
            }
            resolve(result);
          });
        });
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          effectiveUrl = siteUrl;
        }
      }

      let driveIdFromUrl: string | null = null;
      let fileIdFromUrl: string | null = null;
      try {
        const url = new URL(effectiveUrl);
        const pathSegments = url.pathname.split('/').filter(segment => segment);
        const siteIndex = pathSegments.indexOf('sites');
        console.log('DEBUG findDocumentLibrary - effectiveUrl:', effectiveUrl);
        console.log('DEBUG findDocumentLibrary - pathSegments:', pathSegments);
        if (siteIndex !== -1 && pathSegments.length > siteIndex + 2) {
          const driveName = pathSegments[siteIndex + 2];
          const relativePath = pathSegments.slice(siteIndex + 3).join('/');
          console.log('DEBUG findDocumentLibrary - driveName:', driveName);
          console.log('DEBUG findDocumentLibrary - relativePath:', relativePath);
          
          const driveNameMappings: { [key: string]: string[] } = {
            'DOCUMENTOS_CONSULADO': ['DOCUMENTOS_CONSULADO', 'documentos_consulado'],
            'DOCUMENTOS_ADMIN_HH': ['DOCUMENTOS_ADMIN_RRHH', 'documentos_admin_rrhh'],
            'DOCUMENTOS_CLIENTES': ['DOCUMENTOS_CLIENTES', 'documentos_clientes'],
            'DOCUMENTOS_CLIENTES_V2': ['DOCUMENTOS_CLIENTES', 'documentos_clientes'], // Mapear V2 a la versión real
            'DOCUMENTOS_ADMIN_RRHH': ['DOCUMENTOS_ADMIN_RRHH', 'documentos_admin_rrhh'],
          };
          let drive = drives.find((d: DriveInfo) =>
            d.name.toLowerCase() === driveName.toLowerCase() || d.id === driveName
          );
          
          if (!drive && driveNameMappings[driveName]) {
            const possibleNames = driveNameMappings[driveName];
            drive = drives.find((d: DriveInfo) =>
              possibleNames.some(name =>
                d.name === name ||
                d.name.toLowerCase() === name.toLowerCase()
              )
            );
          }

          if (!drive) {
            const driveNameParts = driveName.toLowerCase().split('_');
            drive = drives.find((d: DriveInfo) =>
              driveNameParts.every(part => d.name.toLowerCase().includes(part))
            );
          }
          if (drive) {
            driveIdFromUrl = drive.id;
            const targetFileName = urlDerivedFileName || fileName;
            const decodedRelativePath = decodeURIComponent(relativePath);
            const pathWithNewFileName = decodedRelativePath.replace(/[^/]+$/, targetFileName);

            console.log('DEBUG findDocumentLibrary - drive found:', drive.name);
            console.log('DEBUG findDocumentLibrary - targetFileName:', targetFileName);
            console.log('DEBUG findDocumentLibrary - decodedRelativePath:', decodedRelativePath);
            console.log('DEBUG findDocumentLibrary - pathWithNewFileName:', pathWithNewFileName);

            const fileUrl = `https://graph.microsoft.com/v1.0/drives/${driveIdFromUrl}/root:/${encodeURIComponent(pathWithNewFileName)}`;
            console.log('DEBUG findDocumentLibrary - fileUrl:', fileUrl);
            let attempt = 0;
            while (attempt < MAX_RETRIES) {
              try {
                const fileResult = await getGraphData(fileUrl, accessToken);
                if (fileResult && fileResult.id) {
                  fileIdFromUrl = fileResult.id;
                  return {
                    success: true,
                    data: {
                      driveId: driveIdFromUrl,
                      driveName: drive.name,
                      fileId: fileIdFromUrl,
                      fileName: fileResult.name,
                      file: fileResult,
                      isNewDocument: false
                    }
                  };
                }
                break;
              } catch (fileError) {
                attempt++;
                if (attempt < MAX_RETRIES) {
                  await new Promise(resolve => setTimeout(resolve, RETRY_DELAY * attempt));
                } else {
                  break;
                }
              }
            }
          } else {
            console.log(`Drive ${driveName} no encontrado en la lista de drives disponibles:`, drives.map((d: any) => d.name));
          }
        } else {
          console.log(`URL efectiva no contiene biblioteca válida: ${effectiveUrl}`);
        }
      } catch (urlError) {
        console.log(`Error parseando URL efectiva ${effectiveUrl}:`, urlError.message);
      }

      // If fileId is provided, try direct lookup
      if (fileId) {
        const targetDrives = drives.filter((drive: DriveInfo) =>
          drive.name.toLowerCase() === 'documentos_admin_rrhh' ||
          drive.name.toLowerCase() === 'documentos_clientes'
        );
        for (const drive of targetDrives) {
          let attempt = 0;
          while (attempt < MAX_RETRIES) {
            try {
              const fileUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${fileId}`;
              const fileResult = await getGraphData(fileUrl, accessToken);
              if (fileResult && fileResult.id) {
                return {
                  success: true,
                  data: {
                    driveId: drive.id,
                    driveName: drive.name,
                    fileId: fileResult.id,
                    fileName: fileResult.name,
                    file: fileResult,
                    isNewDocument: false
                  }
                };
              }
              break;
            } catch (fileIdError) {
              attempt++;
              if (attempt < MAX_RETRIES) {
                await new Promise(resolve => setTimeout(resolve, RETRY_DELAY * attempt));
              } else {
                console.log(`Error final buscando por fileId en drive ${drive.name}:`, fileIdError.message);
                break;
              }
            }
          }
        }
      }

      // If URL and fileId fail, mark as new document (no search due to large drives)
      return {
        success: false,
        error: `File "${fileName}" not found via URL or fileId`,
        data: { isNewDocument: true }
      };
    } catch (error) {
      if (retryCount < MAX_RETRIES - 1) {
        await new Promise(resolve => setTimeout(resolve, RETRY_DELAY * (retryCount + 1)));
        return attemptSearch(retryCount + 1);
      }
      console.error('Error final en findDocumentLibrary:', error);
      return {
        success: false,
        error: error.toString(),
        data: { isNewDocument: true }
      };
    } finally {
      if (onLoadingChange) onLoadingChange(false);
    }
  };

  return attemptSearch();
};

export const validateSharePointConnection = async (
  accessToken: string,
  siteUrl: string
): Promise<SharePointResponse> => {
  try {
    const response = await fetch(siteUrl, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json'
      }
    });

    if (!response.ok) { throw new Error(`Site not accessible: ${response.status}`); }
    const data = await response.json();
    return { success: true, data: data };
  } catch (error) {
    console.error('SharePoint validation error:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
};