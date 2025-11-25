import axios, { AxiosResponse, AxiosError } from 'axios';

// Configuración
const CONFIG = {
  timeout: 30000,
  maxRetries: 3,
  retryDelay: 1000
};

// Variable global para almacenar la función de renovación de token
let tokenRenewalFunction: (() => Promise<string | null>) | null = null;

/**
 * Registrar función de renovación de token
 * Esta función debe ser llamada desde office-apis-helpers.ts
 */
export const registerTokenRenewal = (renewalFn: () => Promise<string | null>) => {
  tokenRenewalFunction = renewalFn;
};

// Función para delay
const delay = (ms: number): Promise<void> => 
  new Promise(resolve => setTimeout(resolve, ms));

// Función para obtener delay con backoff exponencial
const getRetryDelay = (attempt: number, retryAfter?: string): number => {
  if (retryAfter) {
    const delay = parseInt(retryAfter) * 1000;
    return Math.min(delay, 60000);
  }
  return Math.min(CONFIG.retryDelay * Math.pow(2, attempt), 30000);
};

// Función para verificar si un error es recuperable
const isRetryableError = (error: AxiosError): boolean => {
  if (!error.response) return true;
  const status = error.response.status;
  return status === 401 || status === 429 || status === 503 || status === 504;
};

// Función para obtener mensaje de error
const getErrorMessage = (error: AxiosError): string => {
  if (error.response) {
    const status = error.response.status;
    switch (status) {
      case 401: return 'Token de acceso inválido o expirado';
      case 403: return 'Acceso denegado - permisos insuficientes';
      case 404: return 'Recurso no encontrado';
      case 429: return 'Demasiadas solicitudes - límite excedido';
      case 503: return 'Servicio no disponible temporalmente';
      case 504: return 'Tiempo de espera agotado';
      default: return `Error ${status}: ${error.response.statusText}`;
    }
  }
  return error.message || 'Error de conexión';
};

/**
 * FUNCIÓN PRINCIPAL para llamadas GET - VERSIÓN SIMPLIFICADA
 */
export const getGraphData = async (
  url: string, 
  accessToken: string, // Token obligatorio
  options: {
    timeout?: number;
    maxRetries?: number;
    headers?: Record<string, string>;
  } = {}
): Promise<AxiosResponse> => {
  
  // Validar que tenemos un token
  if (!accessToken) {
    throw new Error('Token de acceso requerido');
  }

  const {
    timeout = CONFIG.timeout,
    maxRetries = CONFIG.maxRetries,
    headers: additionalHeaders = {}
  } = options;

  let currentToken = accessToken;
  let lastError: AxiosError;
  
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const response = await axios({
        url,
        method: 'get',
        timeout,
        headers: {
          'Authorization': `Bearer ${currentToken}`,
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'X-Request-ID': `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
          ...additionalHeaders
        },
        validateStatus: (status) => status < 500,
        maxRedirects: 5,
        withCredentials: false
      });

      // Si es exitoso
      if (response.status >= 200 && response.status < 300) {
        return response;
      }

      // Si es 401 y tenemos función de renovación
      if (response.status === 401 && attempt < maxRetries && tokenRenewalFunction) {
        try {
          const newToken = await tokenRenewalFunction();
          if (newToken && newToken !== currentToken) {
            currentToken = newToken;
            await delay(getRetryDelay(attempt));
            continue;
          } else {
            console.warn('⚠️ No se pudo renovar el token');
          }
        } catch (renewError) {
          console.error('❌ Error renovando token:', renewError);
        }
      }

      // Para otros errores
      const error = new Error(`HTTP ${response.status}: ${response.statusText}`) as AxiosError;
      error.response = response;
      throw error;

    } catch (error) {
      lastError = error as AxiosError;
      
      const errorMsg = getErrorMessage(lastError);
      console.warn(`⚠️ Error intento ${attempt + 1}: ${errorMsg}`);

      // Si es el último intento o error no recuperable
      if (attempt === maxRetries || !isRetryableError(lastError)) {
        break;
      }

      // Esperar antes del siguiente intento
      const retryAfter = lastError.response?.headers?.['retry-after'];
      const retryDelay = getRetryDelay(attempt, retryAfter);
      
      await delay(retryDelay);
    }
  }

  // Todos los intentos fallaron
  const errorMessage = getErrorMessage(lastError);
  console.error(`❌ Falló después de ${maxRetries + 1} intentos: ${errorMessage}`);
  throw new Error(`Error en Graph API: ${errorMessage}`);
};

/**
 * Función para operaciones POST/PATCH/PUT
 */
export const postGraphData = async (
  url: string,
  data: any,
  accessToken: string, // Token obligatorio
  options: {
    timeout?: number;
    maxRetries?: number;
    method?: 'POST' | 'PATCH' | 'PUT';
    headers?: Record<string, string>;
  } = {}
): Promise<AxiosResponse> => {
  
  if (!accessToken) {
    throw new Error('Token de acceso requerido');
  }

  const {
    timeout = CONFIG.timeout,
    maxRetries = CONFIG.maxRetries,
    method = 'POST',
    headers: additionalHeaders = {}
  } = options;

  let currentToken = accessToken;
  let lastError: AxiosError;
  
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const response = await axios({
        url,
        method: method.toLowerCase() as any,
        data,
        timeout,
        headers: {
          'Authorization': `Bearer ${currentToken}`,
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'X-HTTP-Method': method === 'PATCH' ? 'MERGE' : undefined,
          'X-Request-ID': `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
          ...additionalHeaders
        },
        validateStatus: (status) => status < 500,
        withCredentials: false
      });

      if (response.status >= 200 && response.status < 300) {
        return response;
      }

      if (response.status === 401 && attempt < maxRetries && tokenRenewalFunction) {
        
        try {
          const newToken = await tokenRenewalFunction();
          if (newToken && newToken !== currentToken) {
            currentToken = newToken;
            await delay(getRetryDelay(attempt));
            continue;
          }
        } catch (renewError) {
          console.error('❌ Error renovando token en POST:', renewError);
        }
      }

      const error = new Error(`HTTP ${response.status}: ${response.statusText}`) as AxiosError;
      error.response = response;
      throw error;

    } catch (error) {
      lastError = error as AxiosError;
      
      if (attempt === maxRetries || !isRetryableError(lastError)) {
        break;
      }

      await delay(getRetryDelay(attempt));
    }
  }

  const errorMessage = getErrorMessage(lastError);
  throw new Error(`Error en Graph API ${method}: ${errorMessage}`);
};

/**
 * Función simple para verificar conectividad
 */
export const checkGraphConnectivity = async (accessToken: string): Promise<boolean> => {
  try {
    await getGraphData('https://graph.microsoft.com/v1.0/me', accessToken, {
      timeout: 5000,
      maxRetries: 0
    });
    return true;
  } catch (error) {
    return false;
  }
};