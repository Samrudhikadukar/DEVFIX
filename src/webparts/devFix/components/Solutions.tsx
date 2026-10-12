import * as React from 'react';
import styles from './Solutions.module.scss';

import { SharePointService } from '../services/SharePointService';

import {
  Search,
  ArrowLeft,
  SlidersHorizontal,
  Star,
  ThumbsUp,
  ArrowRight
} from 'lucide-react';

import { WebPartContext } from '@microsoft/sp-webpart-base';

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

export default class Solutions
  extends React.Component<ISolutionsProps, ISolutionsState> {

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

    this.sharePointService =
      new SharePointService(props.context);
  }

  public async componentDidMount(): Promise<void> {

    try {

      const solutions =
        await this.sharePointService.getSolutions();

      this.setState({
        solutions: solutions as ISolution[],
        loading: false
      });

    } catch (error) {

      console.error(
        'Error loading solutions:',
        error
      );

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

    const searchText =
      search.trim().toLowerCase();

    return solutions.filter(solution => {

      if (solution.Status !== 'Published') {
        return false;
      }

      const matchesSearch =
        !searchText ||
        solution.Title
          .toLowerCase()
          .includes(searchText) ||
        solution.Technology
          .toLowerCase()
          .includes(searchText) ||
        solution.Category
          .toLowerCase()
          .includes(searchText) ||
        solution.Problem
          .toLowerCase()
          .includes(searchText);

      const matchesTechnology =
        !technology ||
        solution.Technology === technology;

      const matchesCategory =
        !category ||
        solution.Category === category;

      const matchesDifficulty =
        !difficulty ||
        solution.Difficulty === difficulty;

      return (
        matchesSearch &&
        matchesTechnology &&
        matchesCategory &&
        matchesDifficulty
      );
    });
  }

  private getDifficultyClass(
    difficulty: string
  ): string {

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

  public render(): React.ReactElement<ISolutionsProps> {

    const {
      loading,
      error,
      search,
      technology,
      category,
      difficulty
    } = this.state;

    const filtered =
      this.getFilteredSolutions();

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

    const difficulties = [
      'Easy',
      'Medium',
      'Hard'
    ];

    return (

      <section className={styles.solutionsPage}>

        {/* HEADER */}

        <div className={styles.header}>

          <button
            className={styles.backButton}
            onClick={this.props.onBack}
          >
            <ArrowLeft size={18} />

            Back to Home
          </button>

          <div>

            <h1>
              All Solutions
            </h1>

            <p>
              Browse solutions shared by your
              development team.
            </p>

          </div>

        </div>

        {/* FILTERS */}

        <div className={styles.filterCard}>

          <div className={styles.searchBox}>

            <Search size={18} />

            <input
              type="text"
              placeholder="Search problems, errors, technologies..."
              value={search}
              onChange={event =>
                this.setState({
                  search:
                    event.target.value
                })
              }
            />

          </div>

          <div className={styles.filterTitle}>

            <SlidersHorizontal size={16} />

            Filters

          </div>

          <select
            value={technology}
            onChange={event =>
              this.setState({
                technology:
                  event.target.value
              })
            }
          >

            <option value="">
              All Technologies
            </option>

            {technologies.map(item => (

              <option
                key={item}
                value={item}
              >
                {item}
              </option>

            ))}

          </select>

          <select
            value={category}
            onChange={event =>
              this.setState({
                category:
                  event.target.value
              })
            }
          >

            <option value="">
              All Categories
            </option>

            {categories.map(item => (

              <option
                key={item}
                value={item}
              >
                {item}
              </option>

            ))}

          </select>

          <select
            value={difficulty}
            onChange={event =>
              this.setState({
                difficulty:
                  event.target.value
              })
            }
          >

            <option value="">
              All Difficulties
            </option>

            {difficulties.map(item => (

              <option
                key={item}
                value={item}
              >
                {item}
              </option>

            ))}

          </select>

          <button
            type="button"
            className={styles.clearFilters}
            disabled={
              !search &&
              !technology &&
              !category &&
              !difficulty
            }
            onClick={() =>
              this.setState({
                search: '',
                technology: '',
                category: '',
                difficulty: ''
              })
            }
          >
            ✕ Clear Filters
          </button>

        </div>

        {/* CONTENT */}

        {loading && (

          <div className={styles.message}>
            Loading solutions...
          </div>

        )}

        {error && (

          <div className={styles.message}>
            {error}
          </div>

        )}

        {!loading && !error && (

          <>

            <div className={styles.resultInfo}>

              <strong>
                {filtered.length}
              </strong>

              {' '}solutions found

            </div>

            {filtered.length > 0 ? (

              <div className={styles.solutionGrid}>

                {filtered.map(solution => (

                  <article
                    key={solution.Id}
                    className={styles.solutionCard}
                  >

                    {/* TOP */}

                    <div className={styles.cardTop}>

                      <div
                        className={
                          styles.cardBadges
                        }
                      >

                        <span
                          className={
                            styles.techBadge
                          }
                        >
                          {solution.Technology}
                        </span>

                        <span
                          className={
                            styles.categoryBadge
                          }
                        >
                          {solution.Category}
                        </span>

                        <span
                          className={
                            this.getDifficultyClass(
                              solution.Difficulty
                            )
                          }
                        >
                          {solution.Difficulty}
                        </span>

                      </div>

                      <button
                        type="button"
                        className={
                          styles.favoriteButton
                        }
                        aria-label="Add to favorites"
                        onClick={event =>
                          event.stopPropagation()
                        }
                      >
                        <Star size={16} />
                      </button>

                    </div>

                    {/* TITLE */}

                    <h3>
                      {solution.Title}
                    </h3>

                    {/* PROBLEM */}

                    <p>
                      {solution.Problem}
                    </p>

                    {/* TAGS */}

                    <div
                      className={
                        styles.tags
                      }
                    >

                      {(solution.Tags || '')
                        .split(',')
                        .map(tag =>
                          tag.trim()
                        )
                        .filter(Boolean)
                        .slice(0, 3)
                        .map(tag => (

                          <span
                            key={tag}
                            className={
                              styles.tag
                            }
                          >
                            #{tag}
                          </span>

                        ))}

                    </div>

                    {/* ACTIONS */}

                    <div
                      className={
                        styles.cardActions
                      }
                    >

                      <button
                        type="button"
                        className={
                          styles.helpfulButton
                        }
                        onClick={event =>
                          event.stopPropagation()
                        }
                      >

                        <ThumbsUp size={14} />

                        Helpful ·{' '}
                        {solution.HelpfulCount}

                      </button>

                      <button
  type="button"
  className={styles.viewSolutionButton}
  onClick={event => {
    event.stopPropagation();
    this.props.onViewSolution(solution.Id);
  }}
>
  View Solution
  <ArrowRight size={15} />
</button>

                    </div>

                  </article>

                ))}

              </div>

            ) : (

              <div
                className={
                  styles.emptyState
                }
              >
                No solutions found.
              </div>

            )}

          </>

        )}

      </section>
    );
  }
}