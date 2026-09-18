import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/** File data that the React UI needs to render one training video card. */
export interface ITrainingVideo { title: string; subtitle: string; date: string; duration: string; isNew: boolean; videoUrl: string; }
/** File data that the React UI needs to render one document-library row. */
export interface ITrainingDocument { title: string; type: string; updated: string; documentUrl: string; }
export interface ITrainingLibraryContent { videos: ITrainingVideo[]; documents: ITrainingDocument[]; }
interface ISharePointFile { Name: string; ServerRelativeUrl: string; TimeLastModified: string; }
interface ISharePointFolder { ServerRelativeUrl: string; }
interface ISharePointCollection<T> { value: T[]; }
const VIDEO_EXTENSIONS: string[] = ['mp4', 'webm', 'ogg', 'ogv', 'mov', 'm4v', 'avi', 'wmv'];
const DOCUMENT_EXTENSIONS: string[] = ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv', 'rtf'];

/** Retrieves library content without exposing SharePoint REST details to React UI components. */
export class TrainingLibraryService {
  public constructor(private readonly _spHttpClient: SPHttpClient, private readonly _webAbsoluteUrl: string) {}

  /** The configured library name is resolved to its real root URL; no folder path is hardcoded. */
  public async getLibraryContent(libraryName: string): Promise<ITrainingLibraryContent> {
    const name: string = libraryName.trim();
    if (!name) { return { videos: [], documents: [] }; }
    const files: ISharePointFile[] = await this._getFilesRecursively(await this._getLibraryRootFolderUrl(name));
    const newestFirst: ISharePointFile[] = files.sort((a, b) => new Date(b.TimeLastModified).getTime() - new Date(a.TimeLastModified).getTime());
    return {
      // File extensions separate playable videos from documents returned by the same API.
      videos: newestFirst.filter(file => VIDEO_EXTENSIONS.indexOf(this._getExtension(file.Name)) !== -1).map((file, index) => this._mapVideo(file, index < 3)),
      documents: newestFirst.filter(file => DOCUMENT_EXTENSIONS.indexOf(this._getExtension(file.Name)) !== -1).map(file => this._mapDocument(file))
    };
  }

  /** This REST request finds the configured library and returns its root folder URL. */
  private async _getLibraryRootFolderUrl(libraryName: string): Promise<string> {
    const escapedName: string = libraryName.replace(/'/g, "''");
    const response: SPHttpClientResponse = await this._spHttpClient.get(`${this._webAbsoluteUrl}/_api/web/lists/GetByTitle('${escapedName}')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`, SPHttpClient.configurations.v1);
    if (!response.ok) { throw new Error(`The '${libraryName}' document library could not be found or accessed.`); }
    const library: { RootFolder: ISharePointFolder } = await response.json();
    return library.RootFolder.ServerRelativeUrl;
  }

  /** Retrieves direct files/folders then recurses through every nested folder returned by SharePoint. */
  private async _getFilesRecursively(folderUrl: string): Promise<ISharePointFile[]> {
    const encodedFolderUrl: string = encodeURIComponent(folderUrl).replace(/'/g, '%27');
    const folderApiUrl: string = `${this._webAbsoluteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encodedFolderUrl}')`;
    const [filesResponse, foldersResponse]: [SPHttpClientResponse, SPHttpClientResponse] = await Promise.all([
      this._spHttpClient.get(`${folderApiUrl}/Files?$select=Name,ServerRelativeUrl,TimeLastModified`, SPHttpClient.configurations.v1),
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
  private _mapVideo(file: ISharePointFile, isNew: boolean): ITrainingVideo { return { title: file.Name.replace(/\.[^.]+$/, ''), subtitle: 'Training video', date: `Uploaded ${new Date(file.TimeLastModified).toLocaleDateString()}`, duration: '', isNew, videoUrl: this._getFileUrl(file.ServerRelativeUrl) }; }
  private _mapDocument(file: ISharePointFile): ITrainingDocument { return { title: file.Name.replace(/\.[^.]+$/, ''), type: this._getExtension(file.Name).toUpperCase(), updated: `Updated ${new Date(file.TimeLastModified).toLocaleDateString()}`, documentUrl: this._getFileUrl(file.ServerRelativeUrl) }; }
}
