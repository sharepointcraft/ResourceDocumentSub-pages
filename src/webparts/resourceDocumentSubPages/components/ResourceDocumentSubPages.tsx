import * as React from 'react';
import styles from './ResourceDocumentSubPages.module.scss';
import type { IResourceDocumentSubPagesProps } from './IResourceDocumentSubPagesProps';
import { ITrainingDocument, ITrainingLibraryContent, ITrainingVideo, TrainingLibraryService } from '../services/TrainingLibraryService';

interface IState extends ITrainingLibraryContent {
  isLoading: boolean;
  errorMessage?: string;
  selectedVideo?: ITrainingVideo;
  videoDuration?: string;
}

/** Reusable video card; clicking it opens the in-page player for its SharePoint file. */
const VideoCard: React.FC<{ video: ITrainingVideo; onOpen: (video: ITrainingVideo) => void }> = ({ video, onOpen }) => (
  <article className={styles.videoCard}>
    {/* A button gives the entire card mouse and keyboard support without nesting links. */}
    <button type="button" className={styles.videoCardLink} onClick={() => onOpen(video)} aria-label={`Play video: ${video.title}`}>
    <div className={styles.videoThumbnail}>
      {video.isNew && <span className={styles.newBadge}>New training</span>}
      {video.duration && <span className={styles.duration}>{video.duration}</span>}
      <span className={styles.playIcon} aria-hidden="true">▶</span>
    </div>
    <div className={styles.videoContent}>
      <h3>{video.title}</h3><p>{video.subtitle}</p>
      {video.date && <p className={styles.videoDate}>{video.date}</p>}
      <span className={styles.watchButton}>Watch session</span>
    </div>
    </button>
  </article>
);

/** Main Analyst Training web part component. */
export default class ResourceDocumentSubPages extends React.Component<IResourceDocumentSubPagesProps, IState> {
  private _requestNumber: number = 0;

  public constructor(props: IResourceDocumentSubPagesProps) {
    super(props);
    this.state = { videos: [], documents: [], isLoading: true };
  }

  /** Loads content from the library selected in the property pane when the component appears. */
  public componentDidMount(): void { this._loadLibraryContent().catch(() => undefined); }

  /** A changed property-pane value arrives as a prop and prompts a new dynamic REST request. */
  public componentDidUpdate(previousProps: IResourceDocumentSubPagesProps): void {
    if (previousProps.documentLibraryName !== this.props.documentLibraryName) { this._loadLibraryContent().catch(() => undefined); }
  }

  public componentWillUnmount(): void { this._requestNumber++; }

  /**
   * The service receives the configured library name and authenticated SPFx context. A request
   * number prevents a slow response for an earlier property-pane value overwriting new content.
   */
  private async _loadLibraryContent(): Promise<void> {
    const requestNumber: number = ++this._requestNumber;
    const libraryName: string = this.props.documentLibraryName.trim();
    if (!libraryName) {
      this.setState({ videos: [], documents: [], isLoading: false, errorMessage: undefined });
      return;
    }
    this.setState({ isLoading: true, errorMessage: undefined });
    try {
      const service: TrainingLibraryService = new TrainingLibraryService(this.props.context.spHttpClient, this.props.context.pageContext.web.absoluteUrl);
      const content: ITrainingLibraryContent = await service.getLibraryContent(libraryName);
      if (requestNumber === this._requestNumber) { this.setState({ ...content, isLoading: false }); }
    } catch (error) {
      if (requestNumber === this._requestNumber) {
        this.setState({ videos: [], documents: [], isLoading: false, errorMessage: error instanceof Error ? error.message : 'Training content could not be loaded.' });
      }
    }
  }

  /** Opens the selected video in the viewer and clears duration from a prior video. */
  private _openVideo = (video: ITrainingVideo): void => { this.setState({ selectedVideo: video, videoDuration: undefined }); };

  /** Closes the viewer; unmounting the video element also stops its playback. */
  private _closeVideo = (): void => { this.setState({ selectedVideo: undefined, videoDuration: undefined }); };

  /** Converts the browser-provided duration in seconds into 0:00 or 0:00:00. */
  private _setVideoDuration = (event: React.SyntheticEvent<HTMLVideoElement>): void => {
    const duration: number = event.currentTarget.duration;
    if (!Number.isFinite(duration)) { return; }
    const totalSeconds: number = Math.round(duration);
    // String concatenation keeps this formatter compatible with the SPFx ES target.
    const seconds: string = (`0${totalSeconds % 60}`).slice(-2);
    const minutes: number = Math.floor(totalSeconds / 60) % 60;
    const hours: number = Math.floor(totalSeconds / 3600);
    const paddedMinutes: string = (`0${minutes}`).slice(-2);
    this.setState({ videoDuration: hours > 0 ? `${hours}:${paddedMinutes}:${seconds}` : `${minutes}:${seconds}` });
  };

  public render(): React.ReactElement<IResourceDocumentSubPagesProps> {
    const newestVideos: ITrainingVideo[] = this.state.videos.slice(0, 3);
    const remainingVideos: ITrainingVideo[] = this.state.videos.slice(3);
    const documents: ITrainingDocument[] = this.state.documents;
    return (
      <section className={styles.resourceDocumentSubPages}>
        <header className={styles.hero}><div className={styles.heroInner}>
          <a className={styles.backLink} href="#learning-development">← Back to Learning &amp; Development</a>
          <span className={styles.eyebrow}>Learning collection</span><h1>Analyst Training</h1>
          <p>Build technical excellence and investment insight through curated training content, practical resources, and on-demand sessions designed to accelerate analyst development and long-term success.</p>
        </div></header>
        <main className={styles.content}>
          <div className={styles.mainColumn}>
            {/* Loading and error states replace cards; no static fallback data is displayed. */}
            {this.state.isLoading && <p className={styles.sectionIntro}>Loading training content…</p>}
            {this.state.errorMessage && <p className={styles.sectionIntro} role="alert">{this.state.errorMessage}</p>}
            {/* Empty video sections are omitted; API results alone determine what is shown. */}
            {newestVideos.length > 0 && <section className={styles.videoSection} aria-labelledby="new-videos-title">
              <h2 id="new-videos-title">Newly Uploaded Training Videos</h2>
              <p className={styles.sectionIntro}>The three most recently modified videos in {this.props.documentLibraryName}.</p>
              <div className={styles.videoGrid}>{newestVideos.map(video => <VideoCard key={video.videoUrl} video={video} onOpen={this._openVideo} />)}</div>
            </section>}
            {remainingVideos.length > 0 && <section className={styles.videoSection} aria-labelledby="all-videos-title">
              <h2 id="all-videos-title">All Analyst Training Videos</h2>
              <p className={styles.sectionIntro}>Additional videos from the selected training library.</p>
              <div className={styles.videoGrid}>{remainingVideos.map(video => <VideoCard key={video.videoUrl} video={video} onOpen={this._openVideo} />)}</div>
            </section>}
            <section className={styles.nextTraining}>
              <span className={styles.eyebrow}>Analyst development</span><h2>Next Analyst Training Day</h2>
              <p>Join the panel for confirmed upcoming information, venue details, or a message that new information will be shared soon.</p>
              <a className={styles.calendarButton} href="#calendar">View learning calendar</a>
            </section>
          </div>
          {/* The full document panel is hidden when the configured library contains no supported documents. */}
          {documents.length > 0 && <aside className={styles.documentsPanel} aria-labelledby="documents-title">
            <div className={styles.documentsHeading}><h2 id="documents-title">Analyst Training<br />Documents</h2></div>
            <div className={styles.documentsBody}>
              <div className={styles.documentsLibraryHeader}><h3>Document<br />Library</h3><span className={styles.referenceMaterials}>Reference<br />Materials</span></div>
              {documents.map(document => (
                /* REST provides documentUrl; the link opens the authorized SharePoint file in a new tab. */
                <a className={styles.documentItem} href={document.documentUrl} target="_blank" rel="noopener noreferrer" key={document.documentUrl}>
                  <span className={styles.documentIcon}>◰</span><span><strong>{document.title}</strong><small>{document.updated}</small></span><em>{document.type}</em>
                </a>
              ))}
            </div>
          </aside>}
        </main>
        {/* The player only mounts after a user selects a video, avoiding background downloads. */}
        {this.state.selectedVideo && <div className={styles.videoModal} role="dialog" aria-modal="true" aria-labelledby="video-viewer-title">
          <div className={styles.videoViewer}>
            <div className={styles.videoViewerHeader}>
              <div><span className={styles.viewerEyebrow}>Analyst Training</span><h2 id="video-viewer-title">{this.state.selectedVideo.title}</h2></div>
              <button type="button" className={styles.closeViewer} onClick={this._closeVideo} aria-label="Close video viewer">×</button>
            </div>
            {/* Native controls handle play, seek, sound, fullscreen, and accessible video playback. */}
            <video className={styles.videoPlayer} src={this.state.selectedVideo.videoUrl} controls autoPlay preload="metadata" onLoadedMetadata={this._setVideoDuration}>
              Your browser does not support video playback.
            </video>
            <div className={styles.videoMetadata}>
              <span>▣ {this.state.selectedVideo.created}</span><span>◷ {this.state.videoDuration || 'Loading duration…'}</span>
              <span>▱ {this.state.selectedVideo.fileSize}</span><span>▱ {this.state.selectedVideo.fileType}</span>
            </div>
          </div>
        </div>}
      </section>
    );
  }
}
