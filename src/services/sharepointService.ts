/**
 * SharePoint REST API helper for File and Folder operations
 */

declare global {
  interface Window {
    _spPageContextInfo?: any;
  }
}

export type SpResult<T> =
  | { status: true; data: T; message?: string }
  | { status: false; message: string; error?: any };

function getContext() {
  const ctx = window._spPageContextInfo;
  if (!ctx) {
    // Para desenvolvimento local fora do SharePoint, retornamos um erro claro
    throw new Error(
      'SharePoint context (_spPageContextInfo) não encontrado. Este app deve rodar dentro de uma página SharePoint.'
    );
  }
  return ctx;
}

export function hasSpContext(): boolean {
  return !!window._spPageContextInfo;
}

export function spSiteUrl(): string {
  return getContext().siteAbsoluteUrl;
}

export function spWebRelUrl(): string {
  return getContext().webServerRelativeUrl;
}

// Obtém o path da pasta atual onde o app está rodando
export function getCurrentFolderPath(): string {
  const ctx = getContext();
  const path = ctx.serverRequestPath || '';
  return path.substring(0, path.lastIndexOf('/'));
}

async function parseSpError(resp: Response): Promise<string> {
  try {
    const data = await resp.json();
    return (
      data?.error?.message?.value ||
      data?.odata?.error?.message?.value ||
      resp.statusText ||
      'Erro SharePoint'
    );
  } catch {
    return resp.statusText || 'Erro SharePoint';
  }
}

async function refreshDigest(): Promise<string> {
  const ctx = getContext();
  // No SharePoint clássico muitas vezes o digest já está na página
  if (ctx.formDigestValue) return ctx.formDigestValue;

  const url = `${spSiteUrl()}/_api/contextinfo`;
  const resp = await fetch(url, {
    method: 'POST',
    headers: {
      Accept: 'application/json; odata=verbose',
      'Content-Type': 'application/json; odata=verbose'
    },
    credentials: 'same-origin'
  });
  if (!resp.ok) throw new Error(await parseSpError(resp));
  const data = await resp.json();
  return data?.d?.GetContextWebInformation?.FormDigestValue;
}

/**
 * Lista os arquivos na pasta atual
 */
export async function listFiles(folderPath: string): Promise<SpResult<any[]>> {
  try {
    const url = `${spSiteUrl()}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length`;
    const resp = await fetch(url, {
      method: 'GET',
      headers: { Accept: 'application/json; odata=verbose' },
      credentials: 'same-origin'
    });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    const data = await resp.json();
    return { status: true, data: data?.d?.results || [] };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

/**
 * Obtém o conteúdo de um arquivo (texto)
 */
export async function getFileContent(fileUrl: string): Promise<SpResult<string>> {
  try {
    const url = `${spSiteUrl()}/_api/web/getFileByServerRelativeUrl('${fileUrl}')/$value`;
    const resp = await fetch(url, {
      method: 'GET',
      credentials: 'same-origin'
    });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    const text = await resp.text();
    return { status: true, data: text };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

/**
 * Salva o conteúdo de um arquivo
 */
export async function saveFile(fileUrl: string, content: string): Promise<SpResult<boolean>> {
  try {
    const digest = await refreshDigest();
    const url = `${spSiteUrl()}/_api/web/getFileByServerRelativeUrl('${fileUrl}')/$value`;
    
    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'X-RequestDigest': digest,
        'X-HTTP-Method': 'PUT'
      },
      body: content,
      credentials: 'same-origin'
    });

    if (!resp.ok) {
        // Se falhar o PUT direto, tentamos via Add em modo overwrite
        const folderPath = fileUrl.substring(0, fileUrl.lastIndexOf('/'));
        const fileName = fileUrl.substring(fileUrl.lastIndexOf('/') + 1);
        const addUrl = `${spSiteUrl()}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/Files/Add(url='${fileName}',overwrite=true)`;
        
        const respAdd = await fetch(addUrl, {
            method: 'POST',
            headers: { 'X-RequestDigest': digest },
            body: content,
            credentials: 'same-origin'
        });
        
        if (!respAdd.ok) return { status: false, message: await parseSpError(respAdd) };
    }
    
    return { status: true, data: true };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}
