import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/** File data that the React UI needs to render one training video card. */
export interface ITrainingVideo {
  title: string;
  subtitle: string;
  date: string;
  duration: string;
  isNew: boolean;
  videoUrl: string;
  created: string;
  fileSize: string;
  fileType: string;
}
/** File data that the React UI needs to render one document-library row. */
export interface ITrainingDocument { title: string; type: string; updated: string; documentUrl: string; }
export interface ITrainingLibraryContent { videos: ITrainingVideo[]; documents: ITrainingDocument[]; libraryLastModified?: string; }
interface ISharePointFile { Name: string; ServerRelativeUrl: string; TimeLastModified: string; TimeCreated: string; Length: string; }
interface ISharePointFolder { Name: string; ServerRelativeUrl: string; }
interface ISharePointCollection<T> { value: T[]; }
interface ISharePointDocumentLibrary { Id: string; Title: string; RootFolder: ISharePointFolder; LastItemModifiedDate?: string; }
interface ISharePointLibraryChange { LastItemModifiedDate?: string; }
const VIDEO_EXTENSIONS: string[] = ['mp4', 'webm', 'ogg', 'ogv', 'mov', 'm4v', 'avi', 'wmv'];
const DOCUMENT_EXTENSIONS: string[] = ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv', 'rtf'];

/** Retrieves library content without exposing SharePoint REST details to React UI components. */
export class TrainingLibraryService {
  // Keeps the resolved list ID in memory so polling does not enumerate every site library each time.
  private readonly _resolvedLibraries: { [normalizedConfiguredName: string]: ISharePointDocumentLibrary } = {};

  public constructor(private readonly _spHttpClient: SPHttpClient, private readonly _webAbsoluteUrl: string) {}

  /** The configured library name is resolved to its real root URL; no folder path is hardcoded. */
  public async getLibraryContent(libraryName: string): Promise<ITrainingLibraryContent> {
    const name: string = libraryName.trim();
    if (!name) { return { videos: [], documents: [] }; }
    const library: ISharePointDocumentLibrary = await this._resolveDocumentLibrary(name);
    // All file and folder calls use SharePoint's resolved root URL, never the configured value.
    const [files, libraryLastModified]: [ISharePointFile[], string | undefined] = await Promise.all([
      this._getFilesRecursively(library.RootFolder.ServerRelativeUrl),
      // Reads the current timestamp instead of the cached resolution's original value.
      this._getLibraryLastModifiedById(library.Id)
    ]);
    const videoFiles: ISharePointFile[] = files.filter(file => VIDEO_EXTENSIONS.indexOf(this._getExtension(file.Name)) !== -1);
    // TimeCreated is SharePoint's file Created value. It alone determines the three newly uploaded videos.
    const newestCreatedVideos: ISharePointFile[] = videoFiles.slice().sort((a, b) => new Date(b.TimeCreated).getTime() - new Date(a.TimeCreated).getTime());
    const newVideoFiles: ISharePointFile[] = newestCreatedVideos.slice(0, 3);
    const newVideoUrls: string[] = newVideoFiles.map(file => file.ServerRelativeUrl);
    // The remaining video section retains its existing last-modified ordering.
    const remainingVideos: ISharePointFile[] = videoFiles
      .filter(file => newVideoUrls.indexOf(file.ServerRelativeUrl) === -1)
      .sort((a, b) => new Date(b.TimeLastModified).getTime() - new Date(a.TimeLastModified).getTime());
    const newestFirst: ISharePointFile[] = files.slice().sort((a, b) => new Date(b.TimeLastModified).getTime() - new Date(a.TimeLastModified).getTime());
    return {
      // Created-date videos are first so the unchanged component slice renders them in the new-video section.
      videos: newVideoFiles.map(file => this._mapVideo(file, true)).concat(remainingVideos.map(file => this._mapVideo(file, false))),
      documents: newestFirst.filter(file => DOCUMENT_EXTENSIONS.indexOf(this._getExtension(file.Name)) !== -1).map(file => this._mapDocument(file)),
      // The list-level timestamp is retained for the component's lightweight change detection.
      libraryLastModified
    };
  }

  /** Gets only the list change timestamp; callers reload file data only when this value changes. */
  public async getLibraryLastModified(libraryName: string): Promise<string | undefined> {
    const library: ISharePointDocumentLibrary = await this._resolveDocumentLibrary(libraryName.trim());
    return this._getLibraryLastModifiedById(library.Id);
  }

  /** Reads the current list timestamp by ID after library name resolution has already completed. */
  private async _getLibraryLastModifiedById(libraryId: string): Promise<string | undefined> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      `${this._webAbsoluteUrl}/_api/web/lists(guid'${libraryId}')?$select=LastItemModifiedDate`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) { throw new Error('The configured Learning & Development document library could not be checked for changes.'); }
    const change: ISharePointLibraryChange = await response.json();
    return change.LastItemModifiedDate;
  }

  /** Resolves a configured display name, root-folder name, or root URL to one document library. */
  private async _resolveDocumentLibrary(configuredName: string): Promise<ISharePointDocumentLibrary> {
    const normalizedConfiguredName: string = this._normalizeLibraryValue(configuredName);
    const cachedLibrary: ISharePointDocumentLibrary | undefined = this._resolvedLibraries[normalizedConfiguredName];
    if (cachedLibrary) { return cachedLibrary; }

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      `${this._webAbsoluteUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,LastItemModifiedDate,RootFolder/Name,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=BaseTemplate eq 101`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) { throw new Error('Unable to find the configured Learning & Development document library. Please verify the library display name or internal name.'); }

    const libraries: ISharePointCollection<ISharePointDocumentLibrary> = await response.json();
    const library: ISharePointDocumentLibrary | undefined = libraries.value.find(candidate =>
      !!candidate.RootFolder && [candidate.Title, candidate.RootFolder.Name, candidate.RootFolder.ServerRelativeUrl]
        .some(value => this._normalizeLibraryValue(value) === normalizedConfiguredName)
    );
    if (!library) { throw new Error('Unable to find the configured Learning & Development document library. Please verify the library display name or internal name.'); }

    console.log('[TrainingLibraryService] Resolved document library', {
      configuredValue: configuredName,
      id: library.Id,
      title: library.Title,
      rootFolderName: library.RootFolder.Name,
      rootFolderServerRelativeUrl: library.RootFolder.ServerRelativeUrl
    });
    // Cache only the resolution; the polling method still requests the current change timestamp.
    this._resolvedLibraries[normalizedConfiguredName] = library;
    return library;
  }

  private _normalizeLibraryValue(value: string): string { return (value || '').trim().toLowerCase(); }

  /** Retrieves direct files/folders then recurses through every nested folder returned by SharePoint. */
  private async _getFilesRecursively(folderUrl: string): Promise<ISharePointFile[]> {
    const encodedFolderUrl: string = encodeURIComponent(folderUrl).replace(/'/g, '%27');
    const folderApiUrl: string = `${this._webAbsoluteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encodedFolderUrl}')`;
    const [filesResponse, foldersResponse]: [SPHttpClientResponse, SPHttpClientResponse] = await Promise.all([
      // TimeCreated and Length supply the viewer's dynamic file metadata.
      this._spHttpClient.get(`${folderApiUrl}/Files?$select=Name,ServerRelativeUrl,TimeLastModified,TimeCreated,Length`, SPHttpClient.configurations.v1),
      this._spHttpClient.get(`${folderApiUrl}/Folders?$select=ServerRelativeUrl`, SPHttpClient.configurations.v1)
    ]);
    if (!filesResponse.ok || !foldersResponse.ok) { throw new Error('Files could not be retrieved from one or more library folders.'); }
    const files: ISharePointCollection<ISharePointFile> = await filesResponse.json();
    const folders: ISharePointCollection<ISharePointFolder> = await foldersResponse.json();
    const nestedFiles: ISharePointFile[][] = await Promise.all(folders.value.map(folder => this._getFilesRecursively(folder.ServerRelativeUrl)));
    return files.value.concat(...nestedFiles);
  }

  private _getExtension(name: string): string { const dot: number = name.lastIndexOf('.'); return dot === -1 ? '' : name.substring(dot + 1).toLowerCase(); }
  /** The server-relative URL from REST becomes the real file URL used by UI links. */
  private _getFileUrl(serverRelativeUrl: string): string { return new URL(serverRelativeUrl, this._webAbsoluteUrl).toString(); }
  /** Maps SharePoint file fields into display-ready metadata for the video card and viewer. */
  private _mapVideo(file: ISharePointFile, isNew: boolean): ITrainingVideo {
    const created: Date = new Date(file.TimeCreated);
    const fileType: string = this._getExtension(file.Name).toUpperCase();
    return {
      title: file.Name.replace(/\.[^.]+$/, ''), subtitle: 'Training video', date: `Created ${created.toLocaleDateString()}`,
      duration: '', isNew, videoUrl: this._getFileUrl(file.ServerRelativeUrl),
      created: created.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }),
      fileSize: this._formatFileSize(Number(file.Length)), fileType
    };
  }
  private _mapDocument(file: ISharePointFile): ITrainingDocument { return { title: file.Name.replace(/\.[^.]+$/, ''), type: this._getExtension(file.Name).toUpperCase(), updated: `Updated ${new Date(file.TimeLastModified).toLocaleDateString()}`, documentUrl: this._getFileUrl(file.ServerRelativeUrl) }; }
  /** Converts SharePoint's byte count into a compact, readable size such as 2.2 GB. */
  private _formatFileSize(bytes: number): string {
    if (!Number.isFinite(bytes) || bytes <= 0) { return 'Size unavailable'; }
    const units: string[] = ['B', 'KB', 'MB', 'GB', 'TB'];
    const index: number = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const size: number = bytes / Math.pow(1024, index);
    return `${size >= 10 || index === 0 ? Math.round(size) : size.toFixed(1)} ${units[index]}`;
  }
}
