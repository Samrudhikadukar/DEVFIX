
import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  Icon,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn
} from "@fluentui/react";

import { SharePointService } from "../services/SharePointService";
import { UserService, ICurrentUser } from "../services/UserService";

import styles from "./MyContributions.module.scss";

interface IMyContributionsProps {
  context: WebPartContext;
  onBack: () => void;
  onViewSolution: (id: number) => void;
  onAddSolution: () => void;
  onEditSolution: (id: number) => void;
}

interface IMyContributionsState {
  solutions: any[];
  loading: boolean;
  error: string;
  currentUser: ICurrentUser | null;
  sortField: string;
  isSortedDescending: boolean;
}

export default class MyContributions extends React.Component<
  IMyContributionsProps,
  IMyContributionsState
> {
  private sharePointService: SharePointService;
  private userService: UserService;

  constructor(props: IMyContributionsProps) {
    super(props);

    this.sharePointService = new SharePointService(props.context);
    this.userService = new UserService(props.context);

    this.state = {
      solutions: [],
      loading: true,
      error: "",
      currentUser: null,
      sortField: "",
      isSortedDescending: false
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadContributions();
  }

  private async loadContributions(): Promise<void> {
    try {
      this.setState({
        loading: true,
        error: ""
      });

      const currentUser = await this.userService.getCurrentUser();

      const solutions =
        await this.sharePointService.getMyContributions(
          currentUser.id
        );

      this.setState({
        currentUser,
        solutions,
        loading: false
      });
    } catch (error) {
      console.error("Error loading contributions:", error);

      this.setState({
        loading: false,
        error: "Unable to load your contributions."
      });
    }
  }

  private handleAddSolution = (): void => {
    this.props.onAddSolution();
  };

  private handleView = (id: number): void => {
    this.props.onViewSolution(id);
  };

  private handleEdit = (id: number): void => {
    this.props.onEditSolution(id);
  };

  private handleDelete = async (id: number): Promise<void> => {
    const confirmed = window.confirm(
      "Are you sure you want to delete this solution? This action cannot be undone."
    );

    if (!confirmed) {
      return;
    }

    try {
      await this.sharePointService.deleteSolution(id);

      this.setState(prevState => ({
        solutions: prevState.solutions.filter(
          solution => Number(solution.Id) !== id
        ),
        error: ""
      }));
    } catch (error) {
      console.error("Error deleting solution:", error);

      this.setState({
        error: "Unable to delete the solution."
      });
    }
  };

  private getStatusClass(status: string): string {
    switch (status) {
      case "Published":
        return styles.published;
      case "Pending Review":
        return styles.pending;
      case "Draft":
        return styles.draft;
      default:
        return styles.draft;
    }
  }

  private formatDate(dateValue: string): string {
    if (!dateValue) {
      return "";
    }

    return new Date(dateValue).toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric"
    });
  }

  private handleColumnClick = (
    _event?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ): void => {
    if (!column || !column.fieldName) {
      return;
    }

    this.setState(prevState => ({
      sortField: column.fieldName as string,
      isSortedDescending:
        prevState.sortField === column.fieldName
          ? !prevState.isSortedDescending
          : false
    }));
  };

  private getColumns(): IColumn[] {
    const { sortField, isSortedDescending } = this.state;

    return [
      {
        key: "Title",
        name: "Solution",
        fieldName: "Title",
        minWidth: 180,
        maxWidth: 360,
        isResizable: true,
        isSorted: sortField === "Title",
        isSortedDescending,
        onColumnClick: this.handleColumnClick,
        onRender: (item: any) => (
          <span className={styles.solutionTitleCell}>
            {item.Title}
          </span>
        )
      },
      {
        key: "Technology",
        name: "Technology",
        fieldName: "Technology",
        minWidth: 100,
        maxWidth: 180,
        isResizable: true,
        isSorted: sortField === "Technology",
        isSortedDescending,
        onColumnClick: this.handleColumnClick,
        onRender: (item: any) => (
          <span className={styles.tableTechBadge}>
            {item.Technology}
          </span>
        )
      },
      {
        key: "Category",
        name: "Category",
        fieldName: "Category",
        minWidth: 100,
        maxWidth: 220,
        isResizable: true,
        isSorted: sortField === "Category",
        isSortedDescending,
        onColumnClick: this.handleColumnClick,
        onRender: (item: any) => (
          <span className={styles.tableMutedText}>
            {item.Category}
          </span>
        )
      },
      {
        key: "Status",
        name: "Status",
        fieldName: "Status",
        minWidth: 120,
        maxWidth: 190,
        isResizable: true,
        isSorted: sortField === "Status",
        isSortedDescending,
        onColumnClick: this.handleColumnClick,
        onRender: (item: any) => (
          <span className={this.getStatusClass(item.Status)}>
            {item.Status}
          </span>
        )
      },
      {
        key: "Created",
        name: "Created",
        fieldName: "Created",
        minWidth: 110,
        maxWidth: 170,
        isResizable: true,
        isSorted: sortField === "Created",
        isSortedDescending,
        onColumnClick: this.handleColumnClick,
        onRender: (item: any) => (
          <span className={styles.tableMutedText}>
            {this.formatDate(item.Created)}
          </span>
        )
      },
      {
        key: "Actions",
        name: "Actions",
        fieldName: "Actions",
        minWidth: 125,
        maxWidth: 145,
        isResizable: false,
        onRender: (item: any) => (
          <div className={styles.tableActions}>
            <button
              type="button"
              className={styles.tableIconButton}
              title="View"
              aria-label="View solution"
              onClick={() => this.handleView(Number(item.Id))}
            >
              <Icon iconName="View" />
            </button>

            <button
              type="button"
              className={styles.tableIconButton}
              title="Edit"
              aria-label="Edit solution"
              onClick={() => this.handleEdit(Number(item.Id))}
            >
              <Icon iconName="Edit" />
            </button>

            <button
              type="button"
              className={styles.tableIconButton}
              title="Delete"
              aria-label="Delete solution"
              onClick={() => this.handleDelete(Number(item.Id))}
            >
              <Icon iconName="Delete" />
            </button>
          </div>
        )
      }
    ];
  }

  public render(): React.ReactElement {
    const {
      solutions,
      loading,
      error,
      sortField,
      isSortedDescending
    } = this.state;

    const sortedSolutions = sortField
      ? [...solutions].sort((a: any, b: any) => {
          const first = a[sortField];
          const second = b[sortField];

          let comparison = 0;

          if (sortField === "Created") {
            const firstDate = first
              ? new Date(first).getTime()
              : 0;
            const secondDate = second
              ? new Date(second).getTime()
              : 0;

            comparison = firstDate - secondDate;
          } else {
            comparison = String(first || "").localeCompare(
              String(second || ""),
              undefined,
              { numeric: true, sensitivity: "base" }
            );
          }

          return isSortedDescending
            ? -comparison
            : comparison;
        })
      : solutions;

    return (
      <div className={styles.solutionsPage}>
        <div className={styles.contributionsHeader}>
          <div>
            <h1>My Contributions</h1>
            <p>
              {solutions.length} solution
              {solutions.length === 1 ? "" : "s"} shared by you
            </p>
          </div>

          <button
            type="button"
            className={styles.addSolutionButton}
            onClick={this.handleAddSolution}
          >
            <Icon iconName="Add" />
            Add Solution
          </button>
        </div>

        {error && (
          <div className={styles.message} role="alert">
            {error}
          </div>
        )}

        {loading && (
          <div className={styles.message}>
            Loading your contributions...
          </div>
        )}

        {!loading && !error && solutions.length === 0 && (
          <div className={styles.emptyState}>
            <h3>No contributions yet</h3>
            <p>Share your first fix with the team.</p>
          </div>
        )}

        {!loading && solutions.length > 0 && (
          <div className={styles.contributionsTableCard}>
            <DetailsList
              items={sortedSolutions}
              columns={this.getColumns()}
              setKey="myContributions"
              selectionMode={SelectionMode.none}
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              compact={false}
              className={styles.detailsList}
            />
          </div>
        )}
      </div>
    );
  }
}