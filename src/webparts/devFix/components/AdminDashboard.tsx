import * as React from 'react';
import {
  Icon,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  DefaultButton,
  PrimaryButton
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import styles from './AdminDashboard.module.scss';
import { UserService } from '../services/UserService';
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

type TabType = 'All' | 'Published' | 'Pending Review' | 'Draft' | 'Rejected';
type StatusType = 'Published' | 'Pending Review' | 'Draft' | 'Rejected';

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
  let statusClass = styles.statusSelectDraft;

  if (value === 'Published') {
    statusClass = styles.statusSelectPublished;
  } else if (value === 'Pending Review') {
    statusClass = styles.statusSelectPending;
  }

  return (
    <select
      aria-label="Change status"
      value={value}
      onChange={(event: React.ChangeEvent<HTMLSelectElement>) =>
        onChange(event.target.value as StatusType)
      }
      className={`${styles.statusSelect} ${statusClass}`}
    >
      <option value="Published">Published</option>
      <option value="Pending Review">Pending Review</option>
      <option value="Draft">Draft</option>
      <option value="Rejected">Rejected</option>
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
  onManageUsers
}) => {
  const [activeTab, setActiveTab] = React.useState<TabType>('All');
  const [totalUsers, setTotalUsers] = React.useState<number>(0);
  const [sortField, setSortField] = React.useState<string>('');
  const [isSortedDescending, setIsSortedDescending] =
    React.useState<boolean>(false);

  /* LOAD USERS - EXISTING LOGIC */

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

    loadUsers().catch(error => {
      console.error('Error loading users:', error);
    });
  }, [context]);

  const spService = React.useMemo(
    () => new SharePointService(context),
    [context]
  );

  /* STATISTICS */

  const totalSolutions = solutions.length;

  const pendingSolutions = solutions.filter(
    solution => solution.Status === 'Pending Review'
  );

  const publishedSolutions = solutions.filter(
    solution => solution.Status === 'Published'
  );

  const resolvedSolutions = publishedSolutions.length;

  /* TAB FILTER */

  const filteredSolutions = solutions.filter(solution => {
    if (activeTab === 'All') {
      return true;
    }

    return solution.Status === activeTab;
  });

  /* TECHNOLOGY STATISTICS */

  const technologyCounts: { [key: string]: number } = {};

  solutions.forEach(solution => {
    const technology = solution.Technology || 'Other';

    if (technologyCounts[technology]) {
      technologyCounts[technology] += 1;
    } else {
      technologyCounts[technology] = 1;
    }
  });

  const technologyEntries = Object.keys(technologyCounts)
    .map(technology => ({
      technology,
      count: technologyCounts[technology]
    }))
    .sort((a, b) => b.count - a.count);

  const maxTechnologyCount =
    technologyEntries.length > 0 ? technologyEntries[0].count : 1;

  /* DATE FORMAT */

  const formatDate = (dateString: string): string => {
    if (!dateString) {
      return '-';
    }

    const date = new Date(dateString);

    if (isNaN(date.getTime())) {
      return '-';
    }

    return date.toLocaleDateString('en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });
  };

  /* TECHNOLOGY ICON - FLUENT UI */

  const getTechnologyIcon = (technology: string): React.ReactNode => {
    const tech = technology.toLowerCase();
    let iconName = 'Code';

    if (tech.indexOf('spfx') !== -1) {
      iconName = 'Stack';
    } else if (tech.indexOf('sharepoint') !== -1) {
      iconName = 'Share';
    } else if (tech.indexOf('react') !== -1) {
      iconName = 'Processing';
    } else if (tech.indexOf('pnp') !== -1) {
      iconName = 'Database';
    } else if (
      tech.indexOf('typescript') !== -1 ||
      tech.indexOf('javascript') !== -1
    ) {
      iconName = 'Code';
    } else if (
      tech.indexOf('power apps') !== -1 ||
      tech.indexOf('power automate') !== -1
    ) {
      iconName = 'GridView';
    } else if (tech.indexOf('azure') !== -1) {
      iconName = 'Cloud';
    } else if (tech.indexOf('.net') !== -1) {
      iconName = 'Server';
    }

    return <Icon iconName={iconName} styles={{ root: { fontSize: 20 } }} />;
  };

  /* VIEW */

  const handleView = (solution: ISolution): void => {
    onViewSolution(solution.Id);
  };

  /* EDIT */

  const handleEdit = (solution: ISolution): void => {
    onEditSolution(solution.Id);
  };

  /* DELETE - EXISTING LOGIC */

  const handleDelete = async (solution: ISolution): Promise<void> => {
    const confirmed = window.confirm(
      `Are you sure you want to delete "${solution.Title}"?`
    );

    if (!confirmed) {
      return;
    }

    try {
      await spService.deleteSolution(solution.Id);
      alert('Solution deleted successfully.');
      await onDataChanged();
    } catch (error) {
      console.error('Delete failed:', error);
      alert('Failed to delete solution.');
    }
  };

  /* STATUS CHANGE - EXISTING LOGIC */

  const handleStatusChange = async (
    solution: ISolution,
    status: StatusType
  ): Promise<void> => {
    try {
      await spService.updateSolutionStatus(solution.Id, status);
      await onDataChanged();
    } catch (error) {
      console.error('Status change failed:', error);
      alert('Failed to update solution status.');
    }
  };

  /* SORTING - UI ONLY */

  const handleColumnClick = (
    _event?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ): void => {
    if (!column || !column.fieldName) {
      return;
    }

    const clickedField = column.fieldName;

    if (sortField === clickedField) {
      setIsSortedDescending(previous => !previous);
    } else {
      setSortField(clickedField);
      setIsSortedDescending(false);
    }
  };

  const sortSolutions = (items: ISolution[]): ISolution[] => {
    if (!sortField) {
      return items;
    }

    return [...items].sort((first, second) => {
      let firstValue: string | number = '';
      let secondValue: string | number = '';

      switch (sortField) {
        case 'Title':
          firstValue = first.Title || '';
          secondValue = second.Title || '';
          break;

        case 'Technology':
          firstValue = first.Technology || 'Other';
          secondValue = second.Technology || 'Other';
          break;

        case 'AuthorTitle':
          firstValue = first.Author?.Title || 'Unknown';
          secondValue = second.Author?.Title || 'Unknown';
          break;

        case 'Status':
          firstValue = first.Status || '';
          secondValue = second.Status || '';
          break;

        case 'Created':
          firstValue = first.Created
            ? new Date(first.Created).getTime() || 0
            : 0;
          secondValue = second.Created
            ? new Date(second.Created).getTime() || 0
            : 0;
          break;

        default:
          return 0;
      }

      let comparison = 0;

      if (typeof firstValue === 'number' && typeof secondValue === 'number') {
        comparison = firstValue - secondValue;
      } else {
        comparison = String(firstValue).localeCompare(
          String(secondValue),
          undefined,
          { numeric: true, sensitivity: 'base' }
        );
      }

      return isSortedDescending ? -comparison : comparison;
    });
  };

  /* FLUENT UI DETAILS LIST COLUMNS */

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'SOLUTION',
      fieldName: 'Title',
      minWidth: 170,
      maxWidth: 300,
      isResizable: true,
      isSorted: sortField === 'Title',
      isSortedDescending: sortField === 'Title' && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ISolution) => (
        <span className={styles.solutionCell} title={item.Title}>
          {item.Title || 'Untitled'}
        </span>
      )
    },
    {
      key: 'technology',
      name: 'TECHNOLOGY',
      fieldName: 'Technology',
      minWidth: 115,
      maxWidth: 200,
      isResizable: true,
      isSorted: sortField === 'Technology',
      isSortedDescending: sortField === 'Technology' && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ISolution) => (
        <span className={styles.technologyBadge}>
          {item.Technology || 'Other'}
        </span>
      )
    },
    {
      key: 'author',
      name: 'AUTHOR',
      fieldName: 'AuthorTitle',
      minWidth: 110,
      maxWidth: 200,
      isResizable: true,
      isSorted: sortField === 'AuthorTitle',
      isSortedDescending: sortField === 'AuthorTitle' && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ISolution) => (
        <span className={styles.authorName}>
          {item.Author?.Title || 'Unknown'}
        </span>
      )
    },
    {
      key: 'status',
      name: 'STATUS',
      fieldName: 'Status',
      minWidth: 145,
      maxWidth: 190,
      isResizable: true,
      isSorted: sortField === 'Status',
      isSortedDescending: sortField === 'Status' && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ISolution) => {
        const status = item.Status;

        const isValidStatus =
          status === 'Published' ||
          status === 'Pending Review' ||
          status === 'Draft' ||
          status === 'Rejected';

        return currentUser.role === 'ADMIN' && isValidStatus ? (
          <StatusSelect
            value={status as StatusType}
            onChange={(newStatus: StatusType) =>
              handleStatusChange(item, newStatus)
            }
          />
        ) : (
          <span
            className={`${styles.statusBadge} ${
              status === 'Published'
                ? styles.statusPublished
                : status === 'Pending Review'
                  ? styles.statusPending
                  : styles.statusDraft
            }`}
          >
            <span className={styles.statusDot}>●</span>
            {status}
          </span>
        );
      }
    },
    {
      key: 'created',
      name: 'CREATED',
      fieldName: 'Created',
      minWidth: 115,
      maxWidth: 150,
      isResizable: true,
      isSorted: sortField === 'Created',
      isSortedDescending: sortField === 'Created' && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (item: ISolution) => (
        <span className={styles.createdCell}>
          {formatDate(item.Created)}
        </span>
      )
    },
    {
      key: 'actions',
      name: 'ACTIONS',
      fieldName: 'Actions',
      minWidth: 115,
      maxWidth: 125,
      isResizable: false,
      onRender: (item: ISolution) => (
        <div className={styles.actionsCell}>
          <button
            type="button"
            title="View"
            aria-label="View solution"
            className={styles.iconButton}
            onClick={() => handleView(item)}
          >
            <Icon iconName="View" />
          </button>

          <button
            type="button"
            title="Edit"
            aria-label="Edit solution"
            className={styles.iconButton}
            onClick={() => handleEdit(item)}
          >
            <Icon iconName="Edit" />
          </button>

          <button
            type="button"
            title="Delete"
            aria-label="Delete solution"
            className={`${styles.iconButton} ${styles.deleteButton}`}
            onClick={() => handleDelete(item)}
          >
            <Icon iconName="Delete" />
          </button>
        </div>
      )
    }
  ];

  /* RENDER DETAILS LIST */

  const renderSolutionsList = (
    items: ISolution[],
    emptyMessage: string,
    listKey: string
  ): React.ReactNode => {
    if (items.length === 0) {
      return <div className={styles.emptyCell}>{emptyMessage}</div>;
    }

    return (
      <DetailsList
        items={sortSolutions(items)}
        columns={columns}
        setKey={listKey}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        isHeaderVisible={true}
        compact={false}
        className={styles.detailsList}
      />
    );
  };

  /* VIEW */

  return (
    <div className={styles.page}>
      {/* HEADER */}

      <div className={styles.headerSection}>
        <div className={styles.headerText}>
          <h1 className={styles.pageTitle}>Admin Dashboard</h1>
          <p className={styles.pageSubtitle}>
            Review submissions and keep the knowledge base healthy
          </p>
        </div>

        <div className={styles.headerButtons}>
  <DefaultButton
    text="All Solutions"
    iconProps={{ iconName: 'KnowledgeArticle' }}
    onClick={onAllSolutions}
    className={styles.fluentButton}
  />

  <PrimaryButton
    text="Manage Users"
    iconProps={{ iconName: 'People' }}
    onClick={onManageUsers}
    className={styles.fluentPrimaryButton}
  />
</div>
      </div>

      {/* STAT CARDS */}

      <div className={styles.statsGrid}>
        <div className={styles.statCard}>
          <div className={`${styles.statIcon} ${styles.totalIcon}`}>
            <Icon iconName="ReadingMode" styles={{ root: { fontSize: 25 } }} />
          </div>
          <div className={styles.statContent}>
            <div className={styles.statLabel}>Total Solutions</div>
            <div className={styles.statNumber}>{totalSolutions}</div>
          </div>
        </div>

        <div className={styles.statCard}>
          <div className={`${styles.statIcon} ${styles.openIcon}`}>
            <Icon iconName="Clock" styles={{ root: { fontSize: 25 } }} />
          </div>
          <div className={styles.statContent}>
            <div className={styles.statLabel}>Open Solutions</div>
            <div className={styles.statNumber}>{pendingSolutions.length}</div>
            <div className={styles.statSmallText}>Pending review</div>
          </div>
        </div>

        <div className={`${styles.statCard} ${styles.resolvedCard}`}>
          <div className={`${styles.statIcon} ${styles.resolvedIcon}`}>
            <Icon iconName="Completed" styles={{ root: { fontSize: 25 } }} />
          </div>
          <div className={styles.statContent}>
            <div className={styles.statLabel}>Resolved Solutions</div>
            <div className={styles.statNumber}>{resolvedSolutions}</div>
            <div className={styles.statSmallText}>Published</div>
          </div>
        </div>

        <div className={styles.statCard}>
          <div className={`${styles.statIcon} ${styles.usersIcon}`}>
            <Icon iconName="People" styles={{ root: { fontSize: 25 } }} />
          </div>
          <div className={styles.statContent}>
            <div className={styles.statLabel}>Total Users</div>
            <div className={styles.statNumber}>{totalUsers}</div>
          </div>
        </div>
      </div>

      {/* PENDING SOLUTIONS */}

      <section className={styles.section}>
        <div className={styles.sectionHeading}>
          <h2 className={styles.sectionTitle}>Pending Solutions</h2>
          <p className={styles.sectionSubtitle}>
            Submissions waiting for approval
          </p>
        </div>

        <div className={styles.tableContainer}>
          {renderSolutionsList(
            pendingSolutions,
            'No pending solutions',
            'pendingSolutions'
          )}
        </div>
      </section>

      {/* ALL SOLUTIONS */}

      <section className={styles.section}>
        <div className={styles.allSolutionsHeader}>
          <h2 className={styles.sectionTitle}>All Solutions</h2>

          <div className={styles.tabs}>
            {(
              [
                'All',
                'Published',
                'Pending Review',
                'Draft',
                'Rejected'
              ] as TabType[]
            ).map(tab => (
              <button
                key={tab}
                type="button"
                onClick={() => setActiveTab(tab)}
                className={`${styles.tab} ${
                  activeTab === tab ? styles.activeTab : ''
                }`}
              >
                {tab}
              </button>
            ))}
          </div>
        </div>

        <div className={styles.tableContainer}>
          {renderSolutionsList(
            filteredSolutions,
            'No solutions found',
            'allSolutions'
          )}
        </div>
      </section>

      {/* TECHNOLOGY STATISTICS */}

      <section className={styles.technologySection}>
        <div>
          <h2 className={styles.sectionTitle}>Technology Statistics</h2>
          <p className={styles.sectionSubtitle}>Solutions per technology</p>
        </div>

        <div className={styles.technologyCard}>
          {technologyEntries.map(entry => {
            const percentage = (entry.count / maxTechnologyCount) * 100;

            return (
              <div key={entry.technology} className={styles.technologyRow}>
                <div className={styles.techNameArea}>
                  <div className={styles.techIcon}>
                    {getTechnologyIcon(entry.technology)}
                  </div>
                  <span className={styles.techName}>{entry.technology}</span>
                </div>

                <div className={styles.progressArea}>
                  <div className={styles.progressTrack}>
                    <div
                      className={styles.progressBar}
                      style={{ width: `${percentage}%` }}
                    />
                  </div>
                  <span className={styles.techCount}>{entry.count}</span>
                </div>
              </div>
            );
          })}
        </div>
      </section>

      {/* FOOTER */}

      <footer className={styles.footer}>
        DevFix · Solve. Learn. Share.
      </footer>
    </div>
  );
};

export default AdminDashboard;