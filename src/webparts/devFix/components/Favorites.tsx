
import * as React from "react";
import {
  Stack,
  Text,
  Spinner,
  MessageBar,
  MessageBarType,
  DefaultButton,
  IconButton,
  Icon
} from "@fluentui/react";

import { SharePointService } from "../services/SharePointService";
import { ICurrentUser } from "../services/UserService";
import styles from "./Favorites.module.scss";

export interface IFavoritesProps {
  sharePointService: SharePointService;
  currentUser: ICurrentUser;
  onViewSolution: (id: number) => void;
  onBack: () => void;
}

export interface IFavoritesState {
  solutions: any[];
  loading: boolean;
  error: string;
}

export default class Favorites extends React.Component<
  IFavoritesProps,
  IFavoritesState
> {
  constructor(props: IFavoritesProps) {
    super(props);

    this.state = {
      solutions: [],
      loading: true,
      error: ""
    };
  }

  public componentDidMount(): void {
    void this.loadFavorites();
  }

  private async loadFavorites(): Promise<void> {
    try {
      this.setState({
        loading: true,
        error: ""
      });

      const solutions =
        await this.props.sharePointService.getFavoriteSolutions(
          this.props.currentUser.id
        );

      this.setState({
        solutions,
        loading: false
      });
    } catch (error) {
      console.error("Error loading favorites:", error);

      this.setState({
        loading: false,
        error: "Unable to load your favorite solutions."
      });
    }
  }

  private removeFavorite = async (
    solutionId: number
  ): Promise<void> => {
    try {
      await this.props.sharePointService.removeFavorite(
        solutionId,
        this.props.currentUser.id
      );

      this.setState(prevState => ({
        solutions: prevState.solutions.filter(
          solution => Number(solution.Id) !== solutionId
        ),
        error: ""
      }));
    } catch (error) {
      console.error("Error removing favorite:", error);

      this.setState({
        error: "Unable to remove this solution from favorites."
      });
    }
  };

  private renderSolutionCard(solution: any): JSX.Element {
    return (
      <Stack
        key={solution.Id}
        className={styles.favoriteSolutionCard}
        tokens={{ childrenGap: 12 }}
      >
        <IconButton
          className={styles.removeFavoriteButton}
          title="Remove from favorites"
          ariaLabel="Remove from favorites"
          iconProps={{ iconName: "FavoriteStarFill" }}
          onClick={() =>
            this.removeFavorite(Number(solution.Id))
          }
        />

        <Text
          variant="small"
          className={styles.technology}
        >
          {solution.Technology || "Technology"}
        </Text>

        <Text
          variant="xLarge"
          className={styles.solutionTitle}
        >
          {solution.Title}
        </Text>

        <Stack
          horizontal
          wrap
          tokens={{ childrenGap: 8 }}
          className={styles.badges}
        >
          <Text
            variant="small"
            className={styles.categoryBadge}
          >
            {solution.Category || "General"}
          </Text>

          <Text
            variant="small"
            className={styles.difficultyBadge}
          >
            {solution.Difficulty || "N/A"}
          </Text>
        </Stack>

        <div className={styles.problemSection}>
          <Text
            variant="small"
            className={styles.problemLabel}
          >
            Problem
          </Text>

          <Text
            variant="medium"
            className={styles.problemText}
          >
            {solution.Problem ||
              "No problem description available."}
          </Text>
        </div>

        <div className={styles.cardFooter}>
          <DefaultButton
            className={styles.viewSolutionButton}
            text="View Solution"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={() =>
              this.props.onViewSolution(Number(solution.Id))
            }
          />
        </div>
      </Stack>
    );
  }

  public render(): React.ReactElement<IFavoritesProps> {
    const { solutions, loading, error } = this.state;

    return (
      <div className={styles.favoritesPage}>
        <Stack
          tokens={{ childrenGap: 24 }}
          className={styles.favoritesContainer}
        >
          <div className={styles.pageHeader}>
            <Text
              variant="xxLarge"
              className={styles.pageTitle}
            >
              Favorites
            </Text>

            <Text
              variant="medium"
              className={styles.pageSubtitle}
            >
              Solutions you've saved for later
            </Text>
          </div>

          {error && (
            <MessageBar
              messageBarType={MessageBarType.error}
            >
              {error}
            </MessageBar>
          )}

          {loading && (
            <Stack
              horizontalAlign="center"
              verticalAlign="center"
              className={styles.loadingState}
            >
              <Spinner label="Loading your favorites..." />
            </Stack>
          )}

          {!loading && !error && solutions.length === 0 && (
            <Stack
              horizontalAlign="center"
              verticalAlign="center"
              tokens={{ childrenGap: 12 }}
              className={styles.emptyState}
            >
              <Icon
                iconName="FavoriteStar"
                className={styles.emptyIcon}
              />

              <Text
                variant="xLarge"
                className={styles.emptyTitle}
              >
                No favorite solutions yet
              </Text>

              <Text
                variant="medium"
                className={styles.pageSubtitle}
              >
                Your saved solutions will appear here.
              </Text>
            </Stack>
          )}

          {!loading && !error && solutions.length > 0 && (
            <div className={styles.favoriteGrid}>
              {solutions.map(solution =>
                this.renderSolutionCard(solution)
              )}
            </div>
          )}
        </Stack>
      </div>
    );
  }
}
