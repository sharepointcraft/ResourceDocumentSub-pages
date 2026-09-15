import * as React from 'react';
import styles from './ResourceDocumentSubPages.module.scss';
import type { IResourceDocumentSubPagesProps } from './IResourceDocumentSubPagesProps';

/** Defines the data required to display one video card. */
interface ITrainingVideo {
  title: string;
  subtitle: string;
  date: string;
  duration: string;
  isNew?: boolean;
  videoUrl: string;
}

/** Defines the data required to display one document-library item. */
interface ITrainingDocument {
  title: string;
  type: string;
  updated: string;
  documentUrl: string;
}

// -----------------------------------------------------------------------------
// DATA ARRAYS
// Replace these sample arrays with data from a SharePoint list or API later.
// The UI automatically creates one card/item for every entry in each array.
// -----------------------------------------------------------------------------

const newlyUploadedVideos: ITrainingVideo[] = [
  {
    title: 'Training video title',
    subtitle: 'Add the company banner, presenter details, and learning outcome.',
    date: 'Duration/Uploaded date',
    duration: '',
    isNew: true,
    videoUrl: '#'
  },
  {
    title: 'Training video title',
    subtitle: 'Add the company banner, presenter details, and learning outcome.',
    date: 'Duration/Uploaded date',
    duration: '',
    isNew: true,
    videoUrl: '#'
  },
  {
    title: 'Training video title',
    subtitle: 'Add the company banner, presenter details, and learning outcome.',
    date: 'Duration/Uploaded date',
    duration: '',
    isNew: true,
    videoUrl: '#'
  }
];

const allTrainingVideos: ITrainingVideo[] = [
  {
    title: 'Leveraging Data (& Intro) — Steve Rosenberg',
    subtitle: 'Analyst Training Day · 4 Oct 2023',
    date: '',
    duration: '48:17',
    videoUrl: '#'
  },
  {
    title: 'Generating Ideas — Chris Hawkins',
    subtitle: 'Analyst Training Day · 4 Oct 2023',
    date: '',
    duration: '48:18',
    videoUrl: '#'
  },
  {
    title: 'Monetizing Ideas — Matt Katten',
    subtitle: 'Analyst Training Day · 4 Oct 2023',
    date: '',
    duration: '47:07',
    videoUrl: '#'
  },
  {
    title: 'Q1 Q&A — Roshan',
    subtitle: 'Analyst Training Day · 4 Oct 2023 · Confidential',
    date: '',
    duration: '44:51',
    videoUrl: '#'
  }
];

const documents: ITrainingDocument[] = [
  {
    title: '2023 Analyst Training Day',
    type: 'PDF',
    updated: 'Spring 2023 Training',
    documentUrl: '#'
  },
  {
    title: '2026 Analyst Training Day',
    type: 'PDF',
    updated: 'Spring 2026 Trainings',
    documentUrl: '#'
  }
];

/**
 * Reusable video card used by both video arrays above.
 * Passing a video object creates its thumbnail, details, and watch link.
 */
const VideoCard: React.FC<{ video: ITrainingVideo }> = ({ video }) => (
  <article className={styles.videoCard}>
    {/* Video thumbnail area with optional "new" badge and duration. */}
    <div className={styles.videoThumbnail}>
      {video.isNew && (
        <span className={styles.newBadge}>New training</span>
      )}

      {video.duration && (
        <span className={styles.duration}>{video.duration}</span>
      )}

      <span className={styles.playIcon} aria-hidden="true">▶</span>
    </div>

    {/* Video title, description, date, and action link. */}
    <div className={styles.videoContent}>
      <h3>{video.title}</h3>
      <p>{video.subtitle}</p>

      {video.date && <p className={styles.videoDate}>{video.date}</p>}

      <a className={styles.watchButton} href={video.videoUrl}>
        Watch session
      </a>
    </div>
  </article>
);

/** Main Analyst Training web part component. */
export default class ResourceDocumentSubPages extends React.Component<IResourceDocumentSubPagesProps> {
  public render(): React.ReactElement<IResourceDocumentSubPagesProps> {
    return (
      <section className={styles.resourceDocumentSubPages}>
        {/* Top banner / page introduction. */}
        <header className={styles.hero}>
          <div className={styles.heroInner}>
            <a className={styles.backLink} href="#learning-development">
              ← Back to Learning &amp; Development
            </a>

            <span className={styles.eyebrow}>Learning collection</span>
            <h1>Analyst Training</h1>

            <p>
              Build technical excellence and investment insight through curated
              training content, practical resources, and on-demand sessions
              designed to accelerate analyst development and long-term success.
            </p>
          </div>
        </header>

        <main className={styles.content}>
          {/* Left side: all video sections and training-day message. */}
          <div className={styles.mainColumn}>
            {/* Cards are dynamically created from newlyUploadedVideos. */}
            <section className={styles.videoSection} aria-labelledby="new-videos-title">
              <h2 id="new-videos-title">Newly Uploaded Training Videos</h2>
              <p className={styles.sectionIntro}>
                Feature the latest upload, training content here.
              </p>

              <div className={styles.videoGrid}>
                {newlyUploadedVideos.map((video, index) => (
                  <VideoCard key={`${video.title}-${index}`} video={video} />
                ))}
              </div>
            </section>

            {/* Cards are dynamically created from allTrainingVideos. */}
            <section className={styles.videoSection} aria-labelledby="all-videos-title">
              <h2 id="all-videos-title">All Analyst Training Videos</h2>
              <p className={styles.sectionIntro}>
                Use this area for the complete, searchable analyst training video library.
              </p>

              <div className={styles.videoGrid}>
                {allTrainingVideos.map((video, index) => (
                  <VideoCard key={`${video.title}-${index}`} video={video} />
                ))}
              </div>
            </section>

            {/* Bottom callout for the next training event. */}
            <section className={styles.nextTraining}>
              <span className={styles.eyebrow}>Analyst development</span>
              <h2>Next Analyst Training Day</h2>
              <p>
                Join the panel for confirmed upcoming information, venue details,
                or a message that new information will be shared soon.
              </p>

              <a className={styles.calendarButton} href="#calendar">
                View learning calendar
              </a>
            </section>
          </div>

          {/* Right side: document cards are dynamically created from documents. */}
          <aside className={styles.documentsPanel} aria-labelledby="documents-title">
            <div className={styles.documentsHeading}>
              <h2 id="documents-title">
                Analyst Training<br />Documents
              </h2>
            </div>

            <div className={styles.documentsBody}>
              {/* Document-library title and reference materials label. */}
              <div className={styles.documentsLibraryHeader}>
                <h3>Document<br />Library</h3>
                <span className={styles.referenceMaterials}>Reference<br />Materials</span>
              </div>

              {documents.map((document, index) => (
                <a
                  className={styles.documentItem}
                  href={document.documentUrl}
                  key={`${document.title}-${index}`}
                >
                  <span className={styles.documentIcon}>▰</span>
                  <span>
                    <strong>{document.title}</strong>
                    <small>{document.updated}</small>
                  </span>
                  <em>{document.type}</em>
                </a>
              ))}
            </div>
          </aside>
        </main>
      </section>
    );
  }
}
