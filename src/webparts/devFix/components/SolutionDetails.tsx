
import * as React from "react";
import styles from "./SolutionDetails.module.scss";
import { UserService, ICurrentUser } from "../services/UserService";

import {
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon
} from "@fluentui/react";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SharePointService } from "../services/SharePointService";

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

interface ISolutionDetailsProps {
  context: WebPartContext;
  solutionId: number;
  onBack: () => void;
}

interface ISolutionDetailsState {
  solution: ISolution | null;
  loading: boolean;
  error: string;
  helpfulLoading: boolean;
  currentUser: ICurrentUser | null;
  favoriteLoading: boolean;
}

export default class SolutionDetails extends React.Component<
  ISolutionDetailsProps,
  ISolutionDetailsState
> {
  private sharePointService: SharePointService;
  private userService: UserService;

  public state: ISolutionDetailsState = {
    solution: null,
    loading: true,
    error: "",
    helpfulLoading: false,
    currentUser: null,
    favoriteLoading: false
  };

  constructor(props: ISolutionDetailsProps) {
    super(props);

    this.sharePointService = new SharePointService(props.context);
    this.userService = new UserService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    try {
      const currentUser = await this.userService.getCurrentUser();

      this.setState({ currentUser }, async () => {
        await this.loadSolution();
      });
    } catch (error) {
      console.error("Error loading current user:", error);
      await this.loadSolution();
    }
  }

  private async loadSolution(): Promise<void> {
    try {
      const solution =
        await this.sharePointService.getSolutionById(
          this.props.solutionId
        );

      let isFavorite = false;

      if (this.state.currentUser) {
        isFavorite = await this.sharePointService.isFavorite(
          this.props.solutionId,
          this.state.currentUser.id
        );
      }

      this.setState({
        solution: {
          ...(solution as ISolution),
          IsFavorite: isFavorite
        },
        loading: false
      });
    } catch (error) {
      console.error("Error loading solution details:", error);

      this.setState({
        loading: false,
        error: "Unable to load solution."
      });
    }
  }

  private async handleHelpful(): Promise<void> {
    if (!this.state.solution || this.state.helpfulLoading) {
      return;
    }

    try {
      this.setState({ helpfulLoading: true });

      await this.sharePointService.markHelpful(
        this.state.solution.Id
      );

      this.setState({
        solution: {
          ...this.state.solution,
          HelpfulCount:
            (this.state.solution.HelpfulCount || 0) + 1
        },
        helpfulLoading: false
      });
    } catch (error) {
      console.error("Error marking solution as helpful:", error);

      this.setState({ helpfulLoading: false });
    }
  }

  private async handleFavorite(): Promise<void> {
    if (
      !this.state.solution ||
      !this.state.currentUser ||
      this.state.favoriteLoading
    ) {
      return;
    }

    try {
      this.setState({ favoriteLoading: true });

      const solutionId = this.state.solution.Id;
      const userId = this.state.currentUser.id;

      if (this.state.solution.IsFavorite) {
        await this.sharePointService.removeFavorite(
          solutionId,
          userId
        );

        this.setState({
          solution: {
            ...this.state.solution,
            IsFavorite: false
          },
          favoriteLoading: false
        });
      } else {
        await this.sharePointService.addFavorite(
          solutionId,
          userId
        );

        this.setState({
          solution: {
            ...this.state.solution,
            IsFavorite: true
          },
          favoriteLoading: false
        });
      }
    } catch (error) {
      console.error("Error updating favorite:", error);

      this.setState({ favoriteLoading: false });
    }
  }

  private getDifficultyClass(difficulty: string): string {
    switch (difficulty) {
      case "Easy":
        return styles.easy;
      case "Medium":
        return styles.medium;
      case "Hard":
        return styles.hard;
      default:
        return "";
    }
  }

  private formatDate(date: string): string {
    if (!date) {
      return "";
    }

    return new Date(date).toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric"
    });
  }

  private renderBadge(
    text: string,
    className: string
  ): React.ReactElement {
    return <span className={className}>{text}</span>;
  }

  private renderBlock(
    iconName: string,
    title: string,
    content: string,
    toneClass: string,
    isCode: boolean = false
  ): React.ReactElement {
    return (
      <section className={styles.block}>
        <h2 className={styles.blockTitle}>
          <span className={`${styles.blockIcon} ${toneClass}`}>
            <Icon
              iconName={iconName}
              styles={{ root: { fontSize: 16 } }}
            />
          </span>
          {title}
        </h2>

        {isCode ? (
          <pre className={styles.errorMessage}>{content}</pre>
        ) : (
          <p className={styles.blockContent}>{content}</p>
        )}
      </section>
    );
  }

  public render(): React.ReactElement<ISolutionDetailsProps> {
    const { solution, loading, error } = this.state;

    if (loading) {
      return (
        <section className={styles.detailsPage}>
          <div className={styles.message}>
            <Spinner
              size={SpinnerSize.medium}
              label="Loading solution..."
            />
          </div>
        </section>
      );
    }

    if (error || !solution) {
      return (
        <section className={styles.detailsPage}>
          <div className={styles.message}>
            <h2>Solution not found</h2>

            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
            >
              {error || "This solution may have been deleted."}
            </MessageBar>

            <DefaultButton
              className={styles.backButton}
              onClick={this.props.onBack}
              text="Back to solutions"
              iconProps={{ iconName: "Back" }}
            />
          </div>
        </section>
      );
    }

    const tags = (solution.Tags || "")
      .split(",")
      .map(tag => tag.trim())
      .filter(Boolean);

    return (
      <section className={styles.detailsPage}>
        <div className={styles.content}>
          <DefaultButton
            className={styles.backLink}
            onClick={this.props.onBack}
            text="All solutions"
            iconProps={{ iconName: "Back" }}
            styles={{
              root: {
                background: "transparent",
                border: "none",
                padding: 0,
                minWidth: 0,
                height: "auto"
              },
              rootHovered: {
                background: "transparent"
              },
              label: {
                fontWeight: 500
              }
            }}
          />

          <header className={styles.header}>
            <div className={styles.badges}>
              {this.renderBadge(
                solution.Technology,
                styles.techBadge
              )}

              {this.renderBadge(
                solution.Category,
                styles.categoryBadge
              )}

              {this.renderBadge(
                solution.Difficulty,
                this.getDifficultyClass(solution.Difficulty)
              )}

              {solution.Status !== "Published" &&
                this.renderBadge(
                  solution.Status,
                  styles.statusBadge
                )}
            </div>

            <h1>{solution.Title}</h1>

            <p className={styles.author}>
              by{" "}
              <strong>
                {solution.Author?.Title || "Unknown"}
              </strong>
              {" · "}
              {this.formatDate(solution.Created)}
            </p>

            <div className={styles.actions}>
              <DefaultButton
                className={styles.helpfulButton}
                onClick={() => this.handleHelpful()}
                disabled={this.state.helpfulLoading}
                styles={{
                  root: {
                    height: "auto",
                    minHeight: 36,
                    padding: "9px 13px"
                  },
                  flexContainer: {
                    gap: 7
                  }
                }}
              >
                <Icon
                  iconName="Like"
                  styles={{ root: { fontSize: 16 } }}
                />
                Helpful · {solution.HelpfulCount || 0}
              </DefaultButton>

              <DefaultButton
                className={styles.favoriteButton}
                onClick={() => this.handleFavorite()}
                disabled={
                  this.state.favoriteLoading ||
                  !this.state.currentUser
                }
                styles={{
                  root: {
                    height: "auto",
                    minHeight: 36,
                    padding: "9px 13px"
                  },
                  flexContainer: {
                    gap: 7
                  }
                }}
              >
                <Icon
                  iconName={
                    solution.IsFavorite
                      ? "FavoriteStarFill"
                      : "FavoriteStar"
                  }
                  styles={{ root: { fontSize: 16 } }}
                />

                {solution.IsFavorite
                  ? "Saved to Favorites"
                  : "Add to Favorites"}
              </DefaultButton>
            </div>
          </header>

          <div className={styles.blocks}>
            {this.renderBlock(
              "Bug",
              "Problem",
              solution.Problem,
              styles.dangerTone
            )}

            {solution.ErrorMessage &&
              this.renderBlock(
                "Warning",
                "Error Message",
                solution.ErrorMessage,
                styles.warningTone,
                true
              )}

            {this.renderBlock(
              "Search",
              "Cause",
              solution.Cause,
              styles.violetTone
            )}

            {this.renderBlock(
              "CheckMark",
              "Solution",
              solution.Solution,
              styles.successTone
            )}

            {solution.LessonLearned &&
              this.renderBlock(
                "Lightbulb",
                "Lesson Learned",
                solution.LessonLearned,
                styles.primaryTone
              )}

            {tags.length > 0 && (
              <div className={styles.tags}>
                {tags.map(tag => (
                  <span key={tag} className={styles.tag}>
                    #{tag}
                  </span>
                ))}
              </div>
            )}
          </div>
        </div>
      </section>
    );
  }
}
