/** Client for the local file bridge server running at http://localhost:3001 */

const BRIDGE = "http://localhost:3001";

export interface FileEntry {
  name: string;
  path: string;
  type: "file" | "dir";
  size: number | null;
  modified: string;
  ext: string | null;
}

export interface DirListing {
  dir: string;
  entries: FileEntry[];
}

export interface FileContent {
  path: string;
  ext: string;
  text: string;
  encoding: string;
}

export interface WriteResult {
  ok: boolean;
  path: string;
  size?: number;
}

export interface HomeDirs {
  home: string;
  documents: string;
  desktop: string;
  downloads: string;
}

async function apiFetch<T>(url: string, options?: RequestInit): Promise<T> {
  const res = await fetch(url, options);
  if (!res.ok) {
    const body = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(body.error || res.statusText);
  }
  return res.json() as Promise<T>;
}

export const bridgeClient = {
  /** Check if the bridge server is running. */
  async isAvailable(): Promise<boolean> {
    try {
      await fetch(`${BRIDGE}/api/health`, { signal: AbortSignal.timeout(1500) });
      return true;
    } catch {
      return false;
    }
  },

  /** Get user home, Documents, Desktop, Downloads paths. */
  async getHomeDirs(): Promise<HomeDirs> {
    return apiFetch<HomeDirs>(`${BRIDGE}/api/home`);
  },

  /** List files and subdirectories in a folder. */
  async listDir(dir: string): Promise<DirListing> {
    return apiFetch<DirListing>(`${BRIDGE}/api/files/list?dir=${encodeURIComponent(dir)}`);
  },

  /** Read a file's text content (.txt, .md, .docx, .xlsx, code files). */
  async readFile(filePath: string): Promise<FileContent> {
    return apiFetch<FileContent>(`${BRIDGE}/api/files/read?path=${encodeURIComponent(filePath)}`);
  },

  /** Write text to a file (create or overwrite). */
  async writeFile(filePath: string, text: string): Promise<WriteResult> {
    return apiFetch<WriteResult>(`${BRIDGE}/api/files/write`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: filePath, text }),
    });
  },

  /** Save text as a new file, auto-named if no name given. */
  async saveAs(text: string, suggestedName?: string, dir?: string): Promise<WriteResult> {
    return apiFetch<WriteResult>(`${BRIDGE}/api/files/save-as`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text, suggestedName, dir }),
    });
  },

  /** Delete a file. */
  async deleteFile(filePath: string): Promise<{ ok: boolean }> {
    return apiFetch(`${BRIDGE}/api/files/delete`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: filePath }),
    });
  },

  /** Rename / move a file. */
  async renameFile(from: string, to: string): Promise<{ ok: boolean }> {
    return apiFetch(`${BRIDGE}/api/files/rename`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ from, to }),
    });
  },
};
