
import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Icon } from "@fluentui/react";

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
      currentUser: null
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

  public render(): React.ReactElement {
    const { solutions, loading, error } = this.state;

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
            <Icon
              iconName="Add"
              styles={{ root: { fontSize: 16 } }}
            />
            Add Solution
          </button>
        </div>

        {error && (
          <div className={styles.message}>
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

            <p>
              Share your first fix with the team.
            </p>
          </div>
        )}

        {!loading && solutions.length > 0 && (
          <div className={styles.contributionsTableCard}>
            <table className={styles.contributionsTable}>
              <thead>
                <tr>
                  <th>Solution</th>
                  <th>Technology</th>
                  <th>Category</th>
                  <th>Status</th>
                  <th>Created</th>
                  <th className={styles.actionsHeader}>
                    Actions
                  </th>
                </tr>
              </thead>

              <tbody>
                {solutions.map(solution => (
                  <tr key={solution.Id}>
                    <td className={styles.solutionTitleCell}>
                      <span>{solution.Title}</span>
                    </td>

                    <td>
                      <span className={styles.tableTechBadge}>
                        {solution.Technology}
                      </span>
                    </td>

                    <td className={styles.tableMutedText}>
                      {solution.Category}
                    </td>

                    <td>
                      <span
                        className={this.getStatusClass(
                          solution.Status
                        )}
                      >
                        {solution.Status}
                      </span>
                    </td>

                    <td className={styles.tableMutedText}>
                      {this.formatDate(solution.Created)}
                    </td>

                    <td>
                      <div className={styles.tableActions}>
                        <button
                          type="button"
                          className={styles.tableIconButton}
                          title="View"
                          aria-label="View solution"
                          onClick={() =>
                            this.handleView(Number(solution.Id))
                          }
                        >
                          <Icon
                            iconName="View"
                            styles={{ root: { fontSize: 16 } }}
                          />
                        </button>

                        <button
                          type="button"
                          className={styles.tableIconButton}
                          title="Edit"
                          aria-label="Edit solution"
                          onClick={() =>
                            this.handleEdit(Number(solution.Id))
                          }
                        >
                          <Icon
                            iconName="Edit"
                            styles={{ root: { fontSize: 16 } }}
                          />
                        </button>

                        <button
                          type="button"
                          className={styles.tableIconButton}
                          title="Delete"
                          aria-label="Delete solution"
                          onClick={() =>
                            this.handleDelete(Number(solution.Id))
                          }
                        >
                          <Icon
                            iconName="Delete"
                            styles={{ root: { fontSize: 16 } }}
                          />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    );
  }
}

