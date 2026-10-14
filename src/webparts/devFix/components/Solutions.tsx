
import * as React from 'react';
import styles from './Solutions.module.scss';

import { SharePointService } from '../services/SharePointService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  DefaultButton,
  PrimaryButton,
  TextField,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  IconButton
} from '@fluentui/react';

interface ISolution {
  Id: number;
  Title: string;
  Technology: string;
  Category: string;
  Difficulty: string;
  Problem: string;
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

interface ISolutionsProps {
  context: WebPartContext;
  onBack: () => void;
  onViewSolution: (id: number) => void;
}

interface ISolutionsState {
  solutions: ISolution[];
  loading: boolean;
  error: string;
  search: string;
  technology: string;
  category: string;
  difficulty: string;
}

export default class Solutions extends React.Component<
  ISolutionsProps,
  ISolutionsState
> {
  private sharePointService: SharePointService;

  public state: ISolutionsState = {
    solutions: [],
    loading: true,
    error: '',
    search: '',
    technology: '',
    category: '',
    difficulty: ''
  };

  constructor(props: ISolutionsProps) {
    super(props);
    this.sharePointService = new SharePointService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    try {
      const solutions =
        await this.sharePointService.getSolutions();

      this.setState({
        solutions: solutions as ISolution[],
        loading: false,
        error: ''
      });
    } catch (error) {
      console.error('Error loading solutions:', error);

      this.setState({
        loading: false,
        error: 'Unable to load solutions.'
      });
    }
  }

  private getFilteredSolutions(): ISolution[] {
    const {
      solutions,
      search,
      technology,
      category,
      difficulty
    } = this.state;

    const searchText = search.trim().toLowerCase();

    return solutions.filter(solution => {
      if (solution.Status !== 'Publish') {
        return false;
      }

      const matchesSearch =
        !searchText ||
        (solution.Title || '').toLowerCase().includes(searchText) ||
        (solution.Technology || '').toLowerCase().includes(searchText) ||
        (solution.Category || '').toLowerCase().includes(searchText) ||
        (solution.Problem || '').toLowerCase().includes(searchText);

      const matchesTechnology =
        !technology || solution.Technology === technology;

      const matchesCategory =
        !category || solution.Category === category;

      const matchesDifficulty =
        !difficulty || solution.Difficulty === difficulty;

      return (
        matchesSearch &&
        matchesTechnology &&
        matchesCategory &&
        matchesDifficulty
      );
    });
  }

  private getDifficultyClass(difficulty: string): string {
    switch (difficulty) {
      case 'Easy':
        return styles.easy;
      case 'Medium':
        return styles.medium;
      case 'Hard':
        return styles.hard;
      default:
        return '';
    }
  }

  private clearFilters = (): void => {
    this.setState({
      search: '',
      technology: '',
      category: '',
      difficulty: ''
    });
  };

  public render(): React.ReactElement<ISolutionsProps> {
    const {
      loading,
      error,
      search,
      technology,
      category,
      difficulty
    } = this.state;

    const filtered = this.getFilteredSolutions();

    const technologies = [
      'React',
      'SPFx',
      'SharePoint',
      'PnPjs',
      'TypeScript',
      'JavaScript',
      'Power Automate',
      'Power Apps',
      '.NET',
      'Azure'
    ];

    const categories = [
      'Debugging',
      'Deployment',
      'API',
      'Authentication',
      'Configuration',
      'Permissions',
      'UI/UX',
      'Performance'
    ];

    const difficulties = ['Easy', 'Medium', 'Hard'];

    const technologyOptions: IDropdownOption[] = [
      { key: '', text: 'All Technologies' },
      ...technologies.map(item => ({
        key: item,
        text: item
      }))
    ];

    const categoryOptions: IDropdownOption[] = [
      { key: '', text: 'All Categories' },
      ...categories.map(item => ({
        key: item,
        text: item
      }))
    ];

    const difficultyOptions: IDropdownOption[] = [
      { key: '', text: 'All Difficulties' },
      ...difficulties.map(item => ({
        key: item,
        text: item
      }))
    ];

    const hasActiveFilters = Boolean(
      search || technology || category || difficulty
    );

    return (
      <section className={styles.solutionsPage}>
        {/* HEADER */}
        <header className={styles.header}>
          <DefaultButton
            className={styles.backButton}
            text="Back to Home"
            iconProps={{ iconName: 'Back' }}
            onClick={this.props.onBack}
          />

          <h1>All Solutions</h1>

          <p>
            Browse solutions shared by your development team.
          </p>
        </header>

        {/* SEARCH AND FILTERS */}
        <div className={styles.filterCard}>
          <div className={styles.searchBox}>
            <TextField
              className={styles.searchField}
              placeholder="Search problems, errors, technologies..."
              value={search}
              onChange={(_, value) =>
                this.setState({ search: value || '' })
              }
              ariaLabel="Search solutions"
            />
          </div>

          <div className={styles.filterTitle}>
            <span>Filter solutions</span>
          </div>

          <Dropdown
            className={styles.filterDropdown}
            placeholder="All Technologies"
            ariaLabel="Filter by technology"
            options={technologyOptions}
            selectedKey={technology}
            onChange={(_, option) =>
              this.setState({
                technology: option ? String(option.key) : ''
              })
            }
          />

          <Dropdown
            className={styles.filterDropdown}
            placeholder="All Categories"
            ariaLabel="Filter by category"
            options={categoryOptions}
            selectedKey={category}
            onChange={(_, option) =>
              this.setState({
                category: option ? String(option.key) : ''
              })
            }
          />

          <Dropdown
            className={styles.filterDropdown}
            placeholder="All Difficulties"
            ariaLabel="Filter by difficulty"
            options={difficultyOptions}
            selectedKey={difficulty}
            onChange={(_, option) =>
              this.setState({
                difficulty: option ? String(option.key) : ''
              })
            }
          />

          <DefaultButton
            className={styles.clearFilters}
            text="Clear Filters"
            iconProps={{ iconName: 'ClearFilter' }}
            disabled={!hasActiveFilters}
            onClick={this.clearFilters}
          />
        </div>

        {/* LOADING */}
        {loading && (
          <div className={styles.message}>
            <Spinner
              size={SpinnerSize.medium}
              label="Loading solutions..."
            />
          </div>
        )}

        {/* ERROR */}
        {!loading && error && (
          <div className={styles.message}>
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
            >
              {error}
            </MessageBar>
          </div>
        )}

        {/* SOLUTIONS */}
        {!loading && !error && (
          <>
            <div className={styles.resultInfo}>
              <strong>{filtered.length}</strong>
              {' '}solutions found
            </div>

            {filtered.length > 0 ? (
              <div className={styles.solutionGrid}>
                {filtered.map(solution => (
                  <article
                    key={solution.Id}
                    className={styles.solutionCard}
                  >
                    {/* TECHNOLOGY, CATEGORY AND DIFFICULTY */}
                    <div className={styles.cardTop}>
                      <div className={styles.cardBadges}>
                        <span className={styles.techBadge}>
                          {solution.Technology || 'Technology'}
                        </span>

                        <span className={styles.categoryBadge}>
                          {solution.Category || 'General'}
                        </span>

                        <span
                          className={this.getDifficultyClass(
                            solution.Difficulty
                          )}
                        >
                          {solution.Difficulty || 'N/A'}
                        </span>
                      </div>

                      <IconButton
                        className={styles.favoriteButton}
                        iconProps={{
                          iconName: solution.IsFavorite
                            ? 'FavoriteStarFill'
                            : 'FavoriteStar'
                        }}
                        title={
                          solution.IsFavorite
                            ? 'Remove from favorites'
                            : 'Add to favorites'
                        }
                        ariaLabel={
                          solution.IsFavorite
                            ? 'Remove from favorites'
                            : 'Add to favorites'
                        }
                        onClick={event => event?.stopPropagation()}
                      />
                    </div>

                    {/* TITLE */}
                    <h3>{solution.Title}</h3>

                    {/* PROBLEM DESCRIPTION */}
                    <p>
                      {solution.Problem ||
                        'No problem description available.'}
                    </p>

                    {/* TAGS */}
                    <div className={styles.tags}>
                      {(solution.Tags || '')
                        .split(',')
                        .map(tag => tag.trim())
                        .filter(Boolean)
                        .slice(0, 3)
                        .map((tag, index) => (
                          <span
                            key={`${solution.Id}-${tag}-${index}`}
                            className={styles.tag}
                          >
                            #{tag}
                          </span>
                        ))}
                    </div>

                    {/* ACTIONS */}
                    <div className={styles.cardActions}>
                      <DefaultButton
                        className={styles.helpfulButton}
                        text={`Helpful · ${solution.HelpfulCount || 0}`}
                        iconProps={{ iconName: 'Like' }}
                        onClick={event => event?.stopPropagation()}
                      />

                      <PrimaryButton
                        className={styles.viewSolutionButton}
                        text="View Solution"
                        iconProps={{ iconName: 'Forward' }}
                        onClick={() =>
                          this.props.onViewSolution(solution.Id)
                        }
                      />
                    </div>
                  </article>
                ))}
              </div>
            ) : (
              <div className={styles.emptyState}>
                <h3>No solutions found</h3>
                <p>
                  Try changing your search or clearing the filters.
                </p>

                {hasActiveFilters && (
                  <DefaultButton
                    text="Clear all filters"
                    iconProps={{ iconName: 'ClearFilter' }}
                    onClick={this.clearFilters}
                  />
                )}
              </div>
            )}
          </>
        )}
      </section>
    );
  }
}