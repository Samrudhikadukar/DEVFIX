
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
    this.loadFavorites();
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
        tokens={{ childrenGap: 12 }}
        styles={{
          root: {
            background: "#ffffff",
            border: "1px solid #e5e7eb",
            borderRadius: 12,
            padding: 20,
            boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
            minHeight: 240,
            position: "relative"
          }
        }}
      >
        <IconButton
          title="Remove from favorites"
          ariaLabel="Remove from favorites"
          iconProps={{ iconName: "FavoriteStarFill" }}
          onClick={() =>
            this.removeFavorite(Number(solution.Id))
          }
          styles={{
            root: {
              position: "absolute",
              top: 12,
              right: 12,
              color: "#f59e0b"
            },
            icon: {
              fontSize: 20
            }
          }}
        />

        <Text
          variant="small"
          styles={{
            root: {
              color: "#6366f1",
              fontWeight: 600,
              paddingRight: 35
            }
          }}
        >
          {solution.Technology || "Technology"}
        </Text>

        <Text
          variant="xLarge"
          styles={{
            root: {
              fontWeight: 600,
              color: "#111827",
              paddingRight: 35
            }
          }}
        >
          {solution.Title}
        </Text>

        <Stack
          horizontal
          wrap
          tokens={{ childrenGap: 8 }}
        >
          <Text
            variant="small"
            styles={{
              root: {
                background: "#f3f4f6",
                padding: "4px 8px",
                borderRadius: 6
              }
            }}
          >
            {solution.Category || "General"}
          </Text>

          <Text
            variant="small"
            styles={{
              root: {
                background: "#f3f4f6",
                padding: "4px 8px",
                borderRadius: 6
              }
            }}
          >
            {solution.Difficulty || "N/A"}
          </Text>
        </Stack>

        <Text
          variant="medium"
          styles={{
            root: {
              color: "#4b5563",
              lineHeight: 1.5
            }
          }}
        >
          {solution.Problem || "No problem description available."}
        </Text>

        <DefaultButton
          text="View Solution"
          onClick={() =>
            this.props.onViewSolution(Number(solution.Id))
          }
          styles={{
            root: {
              borderRadius: 8,
              marginTop: "auto"
            }
          }}
        />
      </Stack>
    );
  }

  public render(): React.ReactElement<IFavoritesProps> {
    const { solutions, loading, error } = this.state;

    return (
      <Stack
        tokens={{ childrenGap: 24 }}
        styles={{
          root: {
            padding: "32px 40px",
            maxWidth: 1400,
            margin: "0 auto",
            width: "100%",
            boxSizing: "border-box"
          }
        }}
      >
        <Stack tokens={{ childrenGap: 6 }}>
          <Text
            variant="xxLarge"
            styles={{
              root: {
                fontWeight: 700,
                color: "#111827"
              }
            }}
          >
            Favorites
          </Text>

          <Text
            variant="medium"
            styles={{
              root: {
                color: "#6b7280"
              }
            }}
          >
            Solutions you've saved for later
          </Text>
        </Stack>

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        )}

        {loading && (
          <Stack
            horizontalAlign="center"
            verticalAlign="center"
            styles={{
              root: {
                minHeight: 250
              }
            }}
          >
            <Spinner label="Loading your favorites..." />
          </Stack>
        )}

        {!loading && !error && solutions.length === 0 && (
          <Stack
            horizontalAlign="center"
            verticalAlign="center"
            tokens={{ childrenGap: 12 }}
            styles={{
              root: {
                minHeight: 220,
                background: "#f9fafb",
                border: "1px dashed #d1d5db",
                borderRadius: 12,
                padding: 32,
                textAlign: "center"
              }
            }}
          >
            <Icon
              iconName="FavoriteStar"
              styles={{
                root: {
                  fontSize: 32,
                  color: "#9ca3af"
                }
              }}
            />

            <Text
              variant="xLarge"
              styles={{
                root: {
                  fontWeight: 600,
                  color: "#374151"
                }
              }}
            >
              No favorite solutions yet
            </Text>

            <Text
              variant="medium"
              styles={{
                root: {
                  color: "#6b7280"
                }
              }}
            >
              Your saved solutions will appear here.
            </Text>
          </Stack>
        )}

        {!loading && solutions.length > 0 && (
          <Stack
            styles={{
              root: {
                display: "grid",
                gridTemplateColumns:
                  "repeat(auto-fill, minmax(280px, 1fr))",
                gap: 20
              }
            }}
          >
            {solutions.map(solution =>
              this.renderSolutionCard(solution)
            )}
          </Stack>
        )}
      </Stack>
    );
  }
}

