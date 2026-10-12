import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import styles from './AdminDashboard.module.scss';
import { UserService } from '../services/UserService';

import {
  BookOpen,
  Clock3,
  CheckCircle2,
  Users,
  Eye,
  Pencil,
  Trash2,
  Layers,
  Database,
  Code2,
  Share2,
  Cloud,
  Grid2X2,
  Server,
  Atom
} from 'lucide-react';

import { SharePointService } from '../services/SharePointService';

/* =====================================================
   INTERFACES
===================================================== */

interface ISolution {
  Id: number;
  Title: string;
  Technology: string;
  Category: string;
  Difficulty: string;
  Problem: string;
  ErrorMessage: string;
  Cause: string;
  Solution: string;
  LessonLearned: string;
  Tags: string;
  HelpfulCount: number;
  IsFavorite: boolean;
  Status: string;
  Created: string;

  Author?: {
    Id: number;
    Title: string;
    EMail: string;
  };
}

interface ICurrentUser {
  id: number;
  name: string;
  email: string;
  initials: string;
  role: 'USER' | 'ADMIN';
}

interface IAdminDashboardProps {
  solutions: ISolution[];
  currentUser: ICurrentUser;
  context: WebPartContext;

  onViewSolution: (id: number) => void;
  onEditSolution: (id: number) => void;
  onDataChanged: () => Promise<void>;
  onAllSolutions: () => void;
  onManageUsers: () => void;
  
}

type TabType =
  | 'All'
  | 'Published'
  | 'Pending Review'
  | 'Draft';

type StatusType =
  | 'Published'
  | 'Pending Review'
  | 'Draft'
  | 'Rejected';

interface ITechnologyEntry {
  technology: string;
  count: number;
}

/* =====================================================
   STATUS SELECT
===================================================== */

interface IStatusSelectProps {
  value: StatusType;
  onChange: (status: StatusType) => void;
}

const StatusSelect: React.FC<IStatusSelectProps> = ({
  value,
  onChange
}) => {

  let statusClass: string =
    styles.statusSelectDraft;

  if (value === 'Published') {
    statusClass =
      styles.statusSelectPublished;
  }

  if (value === 'Pending Review') {
    statusClass =
      styles.statusSelectPending;
  }

  return (
    <select
      aria-label="Change status"
      value={value}
      onChange={(
        event: React.ChangeEvent<HTMLSelectElement>
      ) =>
        onChange(
          event.target.value as StatusType
        )
      }
      className={`${styles.statusSelect} ${statusClass}`}
    >

      <option value="Published">
        Published
      </option>

      <option value="Pending Review">
        Pending Review
      </option>

      <option value="Draft">
        Draft
      </option>

      <option value="Rejected">
        Rejected
      </option>

    </select>
  );
};

/* =====================================================
   COMPONENT
===================================================== */

const AdminDashboard: React.FC<IAdminDashboardProps> = ({
  solutions,
  currentUser,
  context,
  onViewSolution,
  onEditSolution,
  onDataChanged,
    onAllSolutions,
onManageUsers,
}) => {

  const [activeTab, setActiveTab] =
    React.useState<TabType>('All');

    const [totalUsers, setTotalUsers] =
  React.useState<number>(0);

React.useEffect(() => {
  const loadUsers = async (): Promise<void> => {
    try {
      const userService = new UserService(context);
      const data = await userService.getUsers();

      setTotalUsers(data.length);
    } catch (error) {
      console.error('Error loading total users:', error);
    }
  };

  loadUsers().catch((error) => {
    console.error('Error loading users:', error);
  });
}, [context]);

  const spService = React.useMemo(
    () => new SharePointService(context),
    [context]
  );

  /* =====================================================
     STATISTICS
  ===================================================== */

  const totalSolutions: number =
    solutions.length;

  const pendingSolutions: ISolution[] =
    solutions.filter(
      (solution: ISolution) =>
        solution.Status === 'Pending Review'
    );

  const publishedSolutions: ISolution[] =
    solutions.filter(
      (solution: ISolution) =>
        solution.Status === 'Published'
    );

  /*
    In our project "Published" represents
    resolved/approved solutions.
  */
  const resolvedSolutions: number =
    publishedSolutions.length;

  /* =====================================================
     TAB FILTER
  ===================================================== */

  const filteredSolutions: ISolution[] =
    solutions.filter(
      (solution: ISolution) => {

        if (activeTab === 'All') {
          return true;
        }

        return (
          solution.Status === activeTab
        );
      }
    );

  /* =====================================================
     TECHNOLOGY STATISTICS
  ===================================================== */

  const technologyCounts: {
    [key: string]: number;
  } = {};

  solutions.forEach(
    (solution: ISolution) => {

      const technology: string =
        solution.Technology || 'Other';

      if (
        technologyCounts[technology]
      ) {
        technologyCounts[technology] =
          technologyCounts[technology] + 1;
      } else {
        technologyCounts[technology] = 1;
      }
    }
  );

  const technologyEntries: ITechnologyEntry[] =
    Object.keys(technologyCounts)
      .map(
        (
          technology: string
        ): ITechnologyEntry => ({
          technology: technology,
          count:
            technologyCounts[technology]
        })
      )
      .sort(
        (
          a: ITechnologyEntry,
          b: ITechnologyEntry
        ): number =>
          b.count - a.count
      );

  const maxTechnologyCount: number =
    technologyEntries.length > 0
      ? technologyEntries[0].count
      : 1;

  /* =====================================================
     DATE FORMAT
  ===================================================== */

  const formatDate = (
    dateString: string
  ): string => {

    if (!dateString) {
      return '-';
    }

    const date: Date =
      new Date(dateString);

    return date.toLocaleDateString(
      'en-US',
      {
        month: 'short',
        day: 'numeric',
        year: 'numeric'
      }
    );
  };

  /* =====================================================
     TECHNOLOGY ICON
  ===================================================== */

  const getTechnologyIcon = (
    technology: string
  ): React.ReactNode => {

    const tech: string =
      technology.toLowerCase();

    if (
      tech.indexOf('spfx') !== -1
    ) {
      return <Layers size={20} />;
    }

    if (
      tech.indexOf('sharepoint') !== -1
    ) {
      return <Share2 size={20} />;
    }

    if (
      tech.indexOf('react') !== -1
    ) {
      return <Atom size={20} />;
    }

    if (
      tech.indexOf('pnp') !== -1
    ) {
      return <Database size={20} />;
    }

    if (
      tech.indexOf('typescript') !== -1
    ) {
      return <Code2 size={20} />;
    }

    if (
      tech.indexOf('javascript') !== -1
    ) {
      return <Code2 size={20} />;
    }

    if (
      tech.indexOf('power apps') !== -1
    ) {
      return <Grid2X2 size={20} />;
    }

    if (
      tech.indexOf('power automate') !== -1
    ) {
      return <Grid2X2 size={20} />;
    }

    if (
      tech.indexOf('azure') !== -1
    ) {
      return <Cloud size={20} />;
    }

    if (
      tech.indexOf('.net') !== -1
    ) {
      return <Server size={20} />;
    }

    return <Code2 size={20} />;
  };

  /* =====================================================
     VIEW
  ===================================================== */

  const handleView = (
    solution: ISolution
  ): void => {

    onViewSolution(solution.Id);
  };

  /* =====================================================
     EDIT
  ===================================================== */

  const handleEdit = (
    solution: ISolution
  ): void => {

    onEditSolution(solution.Id);
  };

  /* =====================================================
     DELETE
  ===================================================== */

  const handleDelete = async (
    solution: ISolution
  ): Promise<void> => {

    const confirmed: boolean =
      window.confirm(
        `Are you sure you want to delete "${solution.Title}"?`
      );

    if (!confirmed) {
      return;
    }

    try {

      await spService.deleteSolution(
        solution.Id
      );

      alert(
        'Solution deleted successfully.'
      );

      /*
       * Refresh data in parent component.
       * No browser reload required.
       */
      await onDataChanged();

    } catch (error) {

      console.error(
        'Delete failed:',
        error
      );

      alert(
        'Failed to delete solution.'
      );
    }
  };

  /* =====================================================
     STATUS CHANGE
  ===================================================== */

 const handleStatusChange = async (
  solution: ISolution,
  status: StatusType
): Promise<void> => {

  try {

    await spService.updateSolutionStatus(
      solution.Id,
      status
    );

    await onDataChanged();

  } catch (error) {

    console.error(
      "Status change failed:",
      error
    );

    alert(
      "Failed to update solution status."
    );
  }
};

  /* =====================================================
     SOLUTION TABLE ROW
  ===================================================== */

  const renderSolutionRow = (
    solution: ISolution
  ): React.ReactNode => {

    const status: string =
      solution.Status;

    const isValidStatus: boolean =
      status === 'Published' ||
      status === 'Pending Review' ||
      status === 'Draft' ||
      status === 'Rejected';

    return (
      <tr
        key={solution.Id}
        className={styles.tableRow}
      >

        {/* SOLUTION */}

        <td
          className={styles.solutionCell}
          title={solution.Title}
        >
          {solution.Title}
        </td>

        {/* TECHNOLOGY */}

        <td
          className={styles.tableCell}
        >
          <span
            className={
              styles.technologyBadge
            }
          >
            {solution.Technology ||
              'Other'}
          </span>
        </td>

        {/* AUTHOR */}

        <td
          className={styles.tableCell}
        >
          <span
            className={
              styles.authorName
            }
          >
            {solution.Author &&
            solution.Author.Title
              ? solution.Author.Title
              : 'Unknown'}
          </span>
        </td>

        {/* STATUS */}

        <td
          className={styles.tableCell}
        >

          {currentUser.role === 'ADMIN' &&
          isValidStatus ? (

            <StatusSelect
              value={
                status as StatusType
              }
              onChange={(
                newStatus: StatusType
              ) =>
                handleStatusChange(
                  solution,
                  newStatus
                )
              }
            />

          ) : (

            <span
              className={`
                ${styles.statusBadge}
                ${
                  status === 'Published'
                    ? styles.statusPublished
                    : status === 'Pending Review'
                      ? styles.statusPending
                      : styles.statusDraft
                }
              `}
            >

              <span
                className={
                  styles.statusDot
                }
              >
                ●
              </span>

              {status}

            </span>
          )}

        </td>

        {/* CREATED */}

        <td
          className={styles.tableCell}
        >
          {formatDate(
            solution.Created
          )}
        </td>

        {/* ACTIONS */}

        <td
          className={
            styles.actionsCell
          }
        >

          {/* VIEW */}

          <button
            type="button"
            title="View"
            className={
              styles.iconButton
            }
            onClick={() =>
              handleView(solution)
            }
          >
            <Eye size={18} />
          </button>

          {/* EDIT */}

          <button
            type="button"
            title="Edit"
            className={
              styles.iconButton
            }
            onClick={() =>
              handleEdit(solution)
            }
          >
            <Pencil size={18} />
          </button>

          {/* DELETE */}

          <button
            type="button"
            title="Delete"
            className={`
              ${styles.iconButton}
              ${styles.deleteButton}
            `}
            onClick={() =>
              handleDelete(solution)
            }
          >
            <Trash2 size={18} />
          </button>

        </td>

      </tr>
    );
  };

  /* =====================================================
     RETURN
  ===================================================== */

  return (
    <div className={styles.page}>

      {/* =================================================
          HEADER
      ================================================= */}

      <div
        className={
          styles.headerSection
        }
      >

        <div>

          <h1
            className={
              styles.pageTitle
            }
          >
            Admin Dashboard
          </h1>

          <p
            className={
              styles.pageSubtitle
            }
          >
            Review submissions and keep
            the knowledge base healthy
          </p>

        </div>

        <div
          className={
            styles.headerButtons
          }
        >

          <button
  type="button"
  className={styles.secondaryButton}
  onClick={onAllSolutions}
>
  All Solutions
</button>

          <button
            type="button"
            className={
              styles.primaryButton
            }
            onClick={onManageUsers}
          >
            Manage Users
          </button>

        </div>

      </div>

      {/* =================================================
          STAT CARDS
      ================================================= */}

      <div
        className={
          styles.statsGrid
        }
      >

        {/* TOTAL SOLUTIONS */}

        <div
          className={
            styles.statCard
          }
        >

          <div
            className={`
              ${styles.statIcon}
              ${styles.totalIcon}
            `}
          >
            <BookOpen size={25} />
          </div>

          <div
            className={
              styles.statContent
            }
          >

            <div
              className={
                styles.statLabel
              }
            >
              Total Solutions
            </div>

            <div
              className={
                styles.statNumber
              }
            >
              {totalSolutions}
            </div>

          </div>

        </div>

        {/* OPEN SOLUTIONS */}

        <div
          className={
            styles.statCard
          }
        >

          <div
            className={`
              ${styles.statIcon}
              ${styles.openIcon}
            `}
          >
            <Clock3 size={25} />
          </div>

          <div
            className={
              styles.statContent
            }
          >

            <div
              className={
                styles.statLabel
              }
            >
              Open Solutions
            </div>

            <div
              className={
                styles.statNumber
              }
            >
              {pendingSolutions.length}
            </div>

            <div
              className={
                styles.statSmallText
              }
            >
              Pending review
            </div>

          </div>

        </div>

        {/* RESOLVED SOLUTIONS */}

        <div
          className={`
            ${styles.statCard}
            ${styles.resolvedCard}
          `}
        >

          <div
            className={`
              ${styles.statIcon}
              ${styles.resolvedIcon}
            `}
          >
            <CheckCircle2 size={25} />
          </div>

          <div
            className={
              styles.statContent
            }
          >

            <div
              className={
                styles.statLabel
              }
            >
              Resolved Solutions
            </div>

            <div
              className={
                styles.statNumber
              }
            >
              {resolvedSolutions}
            </div>

            <div
              className={
                styles.statSmallText
              }
            >
              Published
            </div>

          </div>

        </div>

        {/* TOTAL USERS */}

        <div
          className={
            styles.statCard
          }
        >

          <div
            className={`
              ${styles.statIcon}
              ${styles.usersIcon}
            `}
          >
            <Users size={25} />
          </div>

          <div
            className={
              styles.statContent
            }
          >

            <div
              className={
                styles.statLabel
              }
            >
              Total Users
            </div>

            <div
              className={
                styles.statNumber}
            >
              {totalUsers}
            </div>

          </div>

        </div>

      </div>

      {/* =================================================
          PENDING SOLUTIONS
      ================================================= */}

      <section
        className={
          styles.section
        }
      >

        <div
          className={
            styles.sectionHeading
          }
        >

          <h2
            className={
              styles.sectionTitle
            }
          >
            Pending Solutions
          </h2>

          <p
            className={
              styles.sectionSubtitle
            }
          >
            Submissions waiting for approval
          </p>

        </div>

        <div
          className={
            styles.tableContainer
          }
        >

          <table
            className={
              styles.table
            }
          >

            <thead>

              <tr>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  SOLUTION
                </th>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  TECHNOLOGY
                </th>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  AUTHOR
                </th>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  STATUS
                </th>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  CREATED
                </th>

                <th
                  className={
                    styles.tableHeaderRight
                  }
                >
                  ACTIONS
                </th>

              </tr>

            </thead>

            <tbody>

              {pendingSolutions.length ===
              0 ? (

                <tr>

                  <td
                    colSpan={6}
                    className={
                      styles.emptyCell
                    }
                  >
                    No pending solutions
                  </td>

                </tr>

              ) : (

                pendingSolutions.map(
                  (
                    solution: ISolution
                  ) =>
                    renderSolutionRow(
                      solution
                    )
                )

              )}

            </tbody>

          </table>

        </div>

      </section>

      {/* =================================================
          ALL SOLUTIONS
      ================================================= */}

      <section
        className={
          styles.section
        }
      >

        <div
          className={
            styles.allSolutionsHeader
          }
        >

          <h2
            className={
              styles.sectionTitle
            }
          >
            All Solutions
          </h2>

          <div
            className={
              styles.tabs
            }
          >

            {(
              [
                'All',
                'Published',
                'Pending Review',
                'Draft',
                  'Rejected'

              ] as TabType[]
            ).map(
              (
                tab: TabType
              ) => (

                <button
                  key={tab}
                  type="button"
                  onClick={() =>
                    setActiveTab(tab)
                  }
                  className={`
                    ${styles.tab}
                    ${
                      activeTab === tab
                        ? styles.activeTab
                        : ''
                    }
                  `}
                >
                  {tab}
                </button>

              )
            )}

          </div>

        </div>

        <div
          className={
            styles.tableContainer
          }
        >

          <table
            className={
              styles.table
            }
          >

            <thead>

              <tr>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  SOLUTION
                </th>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  TECHNOLOGY
                </th>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  AUTHOR
                </th>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  STATUS
                </th>

                <th
                  className={
                    styles.tableHeader
                  }
                >
                  CREATED
                </th>

                <th
                  className={
                    styles.tableHeaderRight
                  }
                >
                  ACTIONS
                </th>

              </tr>

            </thead>

            <tbody>

              {filteredSolutions.length ===
              0 ? (

                <tr>

                  <td
                    colSpan={6}
                    className={
                      styles.emptyCell
                    }
                  >
                    No solutions found
                  </td>

                </tr>

              ) : (

                filteredSolutions.map(
                  (
                    solution: ISolution
                  ) =>
                    renderSolutionRow(
                      solution
                    )

                )

              )}

            </tbody>

          </table>

        </div>

      </section>

      {/* =================================================
          TECHNOLOGY STATISTICS
      ================================================= */}

      <section
        className={
          styles.technologySection
        }
      >

        <div>

          <h2
            className={
              styles.sectionTitle
            }
          >
            Technology Statistics
          </h2>

          <p
            className={
              styles.sectionSubtitle
            }
          >
            Solutions per technology
          </p>

        </div>

        <div
          className={
            styles.technologyCard
          }
        >

          {technologyEntries.map(
            (
              entry: ITechnologyEntry
            ) => {

              const technology: string =
                entry.technology;

              const count: number =
                entry.count;

              const percentage: number =
                (count /
                  maxTechnologyCount) *
                100;

              return (
                <div
                  key={technology}
                  className={
                    styles.technologyRow
                  }
                >

                  <div
                    className={
                      styles.techNameArea
                    }
                  >

                    <div
                      className={
                        styles.techIcon
                      }
                    >
                      {getTechnologyIcon(
                        technology
                      )}
                    </div>

                    <span
                      className={
                        styles.techName
                      }
                    >
                      {technology}
                    </span>

                  </div>

                  <div
                    className={
                      styles.progressArea
                    }
                  >

                    <div
                      className={
                        styles.progressTrack
                      }
                    >

                      <div
                        className={
                          styles.progressBar
                        }
                        style={{
                          width:
                            `${percentage}%`
                        }}
                      />

                    </div>

                    <span
                      className={
                        styles.techCount
                      }
                    >
                      {count}
                    </span>

                  </div>

                </div>
              );
            }
          )}

        </div>

      </section>

      {/* =================================================
          FOOTER
      ================================================= */}

      <footer
        className={
          styles.footer
        }
      >
        DevFix · Solve. Learn. Share.
      </footer>

    </div>
  );
};

export default AdminDashboard;