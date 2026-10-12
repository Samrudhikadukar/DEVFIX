import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Eye, Pencil, Plus, Trash2 } from "lucide-react";

import { SharePointService } from "../services/SharePointService";
import { UserService, ICurrentUser } from "../services/UserService";

//import styles from "./Solutions.module.scss";
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

      this.setState({
        solutions: this.state.solutions.filter(
          solution => solution.Id !== id
        )
      });
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
    const {
      solutions,
      loading,
      error
    } = this.state;

    return (
      <div className={styles.solutionsPage}>

        {/* =========================
            MY CONTRIBUTIONS HEADER
        ========================= */}

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
            <Plus size={16} />
            Add Solution
          </button>

        </div>


        {/* =========================
            ERROR
        ========================= */}

        {error && (
          <div className={styles.message}>
            {error}
          </div>
        )}


        {/* =========================
            LOADING
        ========================= */}

        {loading && (
          <div className={styles.message}>
            Loading your contributions...
          </div>
        )}


        {/* =========================
            EMPTY STATE
        ========================= */}

        {!loading && !error && solutions.length === 0 && (
          <div className={styles.emptyState}>

            <h3>No contributions yet</h3>

            <p>
              Share your first fix with the team.
            </p>

          </div>
        )}


        {/* =========================
            TABLE
        ========================= */}

        {!loading && solutions.length > 0 && (

          <div className={styles.contributionsTableCard}>

            <table className={styles.contributionsTable}>

              <thead>
                <tr>

                  <th>
                    Solution
                  </th>

                  <th>
                    Technology
                  </th>

                  <th>
                    Category
                  </th>

                  <th>
                    Status
                  </th>

                  <th>
                    Created
                  </th>

                  <th className={styles.actionsHeader}>
                    Actions
                  </th>

                </tr>
              </thead>


              <tbody>

                {solutions.map((solution) => (

                  <tr key={solution.Id}>

                    {/* Solution */}

                    <td className={styles.solutionTitleCell}>
                      <span>
                        {solution.Title}
                      </span>
                    </td>


                    {/* Technology */}

                    <td>
                      <span className={styles.tableTechBadge}>
                        {solution.Technology}
                      </span>
                    </td>


                    {/* Category */}

                    <td className={styles.tableMutedText}>
                      {solution.Category}
                    </td>


                    {/* Status */}

                    <td>
                      <span
                        className={this.getStatusClass(
                          solution.Status
                        )}
                      >
                        {solution.Status}
                      </span>
                    </td>


                    {/* Created */}

                    <td className={styles.tableMutedText}>
                      {this.formatDate(
                        solution.Created
                      )}
                    </td>


                    {/* Actions */}

                    <td>

                      <div className={styles.tableActions}>

                        {/* View */}

                        <button
                          type="button"
                          className={styles.tableIconButton}
                          title="View"
                          onClick={() =>
                            this.handleView(solution.Id)
                          }
                        >
                          <Eye size={16} />
                        </button>


                        {/* Edit */}

                        <button
                          type="button"
                          className={styles.tableIconButton}
                          title="Edit"
                          onClick={() =>
                            this.handleEdit(solution.Id)
                          }
                        >
                          <Pencil size={16} />
                        </button>


                        {/* Delete */}

                        <button
                          type="button"
                          className={styles.tableIconButton}
                          title="Delete"
                          onClick={() =>
                            this.handleDelete(solution.Id)
                          }
                        >
                          <Trash2 size={16} />
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