import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import styles from './DevFix.module.scss';
import type { IDevFixProps } from './IDevFixProps';

import { SharePointService } from '../services/SharePointService';
import { UserService } from '../services/UserService';

import DevFixNavbar from './DevFixNavbar';
import Solutions from './Solutions';
import SolutionDetails from './SolutionDetails';
import MyContributions from './MyContributions';
import AddSolutionForm, { IAddSolutionInput } from './AddSolutionForm';
import AdminDashboard from './AdminDashboard';
import ManageUsers from './ManageUsers';
import Favorites from './Favorites';

import {
  BookOpen,
  Layers,
  Sparkles,
  ThumbsUp,
  User,
  ArrowRight
} from 'lucide-react';

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

interface IDevFixState {
  solutions: ISolution[];
  loading: boolean;
  error: string;
  currentUser: ICurrentUser | null;
  search: string;
  technology: string;
  category: string;
  difficulty: string;
  page:
    | 'home'
    | 'solutions'
    | 'favorites'
    | 'details'
    | 'my-contributions'
    | 'add-solution'
    | 'admin'
    | 'manage-users';
  selectedSolutionId: number | null;
  isAdminViewingSolution: boolean;
}

export default class DevFix extends React.Component<
  IDevFixProps,
  IDevFixState
> {
  private sharePointService: SharePointService;
  private userService: UserService;

  public state: IDevFixState = {
    solutions: [],
    loading: true,
    error: '',
    currentUser: null,
    search: '',
    technology: '',
    category: '',
    difficulty: '',
    page: 'home',
    selectedSolutionId: null,
    isAdminViewingSolution: false
  };

  constructor(props: IDevFixProps) {
    super(props);

    this.sharePointService = new SharePointService(props.context);
    this.userService = new UserService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    try {
      const [solutions, currentUser] = await Promise.all([
        this.sharePointService.getSolutions(),
        this.userService.getCurrentUser()
      ]);

      this.setState({
        solutions: solutions as ISolution[],
        currentUser,
        loading: false
      });
    } catch (error) {
      console.error('Error loading DevFix:', error);

      this.setState({
        loading: false,
        error: 'Unable to load DevFix data.'
      });
    }
  }

  private refreshSolutions = async (): Promise<void> => {
    try {
      const solutions = await this.sharePointService.getSolutions();

      this.setState({
        solutions: solutions as ISolution[],
        error: ''
      });
    } catch (error) {
      console.error('Error refreshing solutions:', error);

      this.setState({
        error: 'Unable to refresh solutions.'
      });
    }
  };

  private handleHome = (): void => {
    this.setState({
      page: 'home',
      selectedSolutionId: null,
      isAdminViewingSolution: false
    });
  };

  private handleSolutions = (): void => {
    this.setState({
      page: 'solutions',
      selectedSolutionId: null,
      isAdminViewingSolution: false
    });
  };

  private handleFavorites = (): void => {
    this.setState({
      page: 'favorites',
      selectedSolutionId: null,
      isAdminViewingSolution: false
    });
  };

  private handleMyContributions = (): void => {
    this.setState({
      page: 'my-contributions',
      selectedSolutionId: null,
      isAdminViewingSolution: false
    });
  };

  private handleAdmin = (): void => {
    if (this.state.currentUser?.role !== 'ADMIN') {
      return;
    }

    this.setState({
      page: 'admin',
      selectedSolutionId: null,
      isAdminViewingSolution: false
    });
  };

  private handleManageUsers = (): void => {
    if (this.state.currentUser?.role !== 'ADMIN') {
      return;
    }

    this.setState({
      page: 'manage-users',
      selectedSolutionId: null,
      isAdminViewingSolution: false
    });
  };

  private handleAddSolution = (): void => {
    this.setState({
      page: 'add-solution',
      selectedSolutionId: null,
      isAdminViewingSolution: false
    });
  };

  private handleAddSolutionSubmit = async (
    input: IAddSolutionInput
  ): Promise<void> => {
    try {
      if (this.state.selectedSolutionId !== null) {
        await this.sharePointService.updateSolution(
          this.state.selectedSolutionId,
          input
        );

        await this.refreshSolutions();

        this.setState({
          page: 'admin',
          selectedSolutionId: null,
          isAdminViewingSolution: false,
          error: ''
        });

        return;
      }

      await this.sharePointService.addSolution({
        ...input,
        status: 'Pending Review'
      });

      await this.refreshSolutions();

      this.setState({
        page: 'my-contributions',
        selectedSolutionId: null,
        isAdminViewingSolution: false,
        error: ''
      });
    } catch (error) {
      console.error('Error saving solution:', error);

      this.setState({
        error: 'Unable to save the solution.'
      });

      throw error;
    }
  };

  private handleAdminViewSolution = (id: number): void => {
    this.setState({
      page: 'details',
      selectedSolutionId: id,
      isAdminViewingSolution: true
    });
  };

  private handleAdminEditSolution = (id: number): void => {
    this.setState({
      page: 'add-solution',
      selectedSolutionId: id,
      isAdminViewingSolution: false
    });
  };

  private handleAdminDataChanged = async (): Promise<void> => {
    await this.refreshSolutions();
  };

  private getPublishedSolutions(): ISolution[] {
    return this.state.solutions.filter(
      solution => solution.Status === 'Published'
    );
  }

  private getFilteredSolutions(): ISolution[] {
    const published = this.getPublishedSolutions();
    const search = this.state.search.trim().toLowerCase();
    const technology = this.state.technology;
    const category = this.state.category;
    const difficulty = this.state.difficulty;

    return published.filter(solution => {
      const matchesSearch =
        !search ||
        solution.Title.toLowerCase().includes(search) ||
        solution.Technology.toLowerCase().includes(search) ||
        solution.Category.toLowerCase().includes(search) ||
        solution.Problem.toLowerCase().includes(search);

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

  private formatDate(date: string): string {
    return new Date(date).toLocaleDateString('en-IN', {
      day: 'numeric',
      month: 'short',
      year: 'numeric'
    });
  }

  private getTechnologyCount(technology: string): number {
    return this.getPublishedSolutions().filter(
      solution => solution.Technology === technology
    ).length;
  }

  public render(): React.ReactElement<IDevFixProps> {
    const {
      loading,
      error,
      currentUser,
      search,
      technology,
      category,
      difficulty,
      page
    } = this.state;

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

    const filtered = this.getFilteredSolutions();
    const published = this.getPublishedSolutions();

    const mostHelpful = [...published]
      .sort((a, b) => b.HelpfulCount - a.HelpfulCount)
      .slice(0, 5);

    const recent = [...published]
      .sort(
        (a, b) =>
          new Date(b.Created).getTime() -
          new Date(a.Created).getTime()
      )
      .slice(0, 5);

    const myContributions = currentUser
      ? this.state.solutions.filter(
          solution => solution.Author?.Id === currentUser.id
        ).length
      : 0;

    if (
      page === 'details' &&
      this.state.selectedSolutionId !== null
    ) {
      return (
        <div className={styles.devFix}>
          {currentUser && (
            <DevFixNavbar
              currentUser={currentUser}
              onHome={this.handleHome}
              onSolutions={this.handleSolutions}
              onFavorites={this.handleFavorites}
              onMyContributions={this.handleMyContributions}
              onAddSolution={this.handleAddSolution}
              onAdmin={this.handleAdmin}
            />
          )}

          <SolutionDetails
            context={this.props.context}
            solutionId={this.state.selectedSolutionId}
            onBack={() =>
              this.setState({
                page: this.state.isAdminViewingSolution
                  ? 'admin'
                  : 'solutions',
                selectedSolutionId: null,
                isAdminViewingSolution: false
              })
            }
          />
        </div>
      );
    }

    if (
      page === 'manage-users' &&
      currentUser?.role === 'ADMIN'
    ) {
      return (
        <div className={styles.devFix}>
          <DevFixNavbar
            currentUser={currentUser}
            onHome={this.handleHome}
            onSolutions={this.handleSolutions}
            onFavorites={this.handleFavorites}
            onMyContributions={this.handleMyContributions}
            onAddSolution={this.handleAddSolution}
            onAdmin={this.handleAdmin}
          />

          <ManageUsers
            context={this.props.context}
            solutions={this.state.solutions}
            onBack={this.handleAdmin}
          />
        </div>
      );
    }

    if (page === 'favorites' && currentUser) {
      return (
        <div className={styles.devFix}>
          <DevFixNavbar
            currentUser={currentUser}
            onHome={this.handleHome}
            onSolutions={this.handleSolutions}
            onFavorites={this.handleFavorites}
            onMyContributions={this.handleMyContributions}
            onAddSolution={this.handleAddSolution}
            onAdmin={this.handleAdmin}
          />

          <Favorites
            sharePointService={this.sharePointService}
            currentUser={currentUser}
            onViewSolution={(id: number) => {
              this.setState({
                page: 'details',
                selectedSolutionId: id,
                isAdminViewingSolution: false
              });
            }}
            onBack={this.handleSolutions}
          />
        </div>
      );
    }

    if (page === 'admin' && currentUser?.role === 'ADMIN') {
      return (
        <div className={styles.devFix}>
          <DevFixNavbar
            currentUser={currentUser}
            onHome={this.handleHome}
            onSolutions={this.handleSolutions}
            onFavorites={this.handleFavorites}
            onMyContributions={this.handleMyContributions}
            onAddSolution={this.handleAddSolution}
            onAdmin={this.handleAdmin}
          />

          <AdminDashboard
            solutions={this.state.solutions}
            currentUser={currentUser}
            context={this.props.context}
            onViewSolution={this.handleAdminViewSolution}
            onEditSolution={this.handleAdminEditSolution}
            onDataChanged={this.handleAdminDataChanged}
            onAllSolutions={this.handleSolutions}
            onManageUsers={this.handleManageUsers}
          />
        </div>
      );
    }

    if (page === 'add-solution') {
      const selectedSolution =
        this.state.selectedSolutionId !== null
          ? this.state.solutions.find(
              solution =>
                solution.Id === this.state.selectedSolutionId
            )
          : undefined;

      const editSolution: IAddSolutionInput | undefined =
        selectedSolution
          ? {
              title: selectedSolution.Title || '',
              technology: selectedSolution.Technology || '',
              category: selectedSolution.Category || '',
              difficulty: selectedSolution.Difficulty || '',
              problem: selectedSolution.Problem || '',
              errorMessage: selectedSolution.ErrorMessage || '',
              cause: selectedSolution.Cause || '',
              solution: selectedSolution.Solution || '',
              lessonLearned: selectedSolution.LessonLearned || '',
              tags: selectedSolution.Tags || ''
            }
          : undefined;

      return (
        <div className={styles.devFix}>
          {currentUser && (
            <DevFixNavbar
              currentUser={currentUser}
              onHome={this.handleHome}
              onSolutions={this.handleSolutions}
              onFavorites={this.handleFavorites}
              onMyContributions={this.handleMyContributions}
              onAddSolution={this.handleAddSolution}
              onAdmin={this.handleAdmin}
            />
          )}

          <AddSolutionForm
            context={this.props.context}
            initialSolution={editSolution}
            isEditMode={this.state.selectedSolutionId !== null}
            onCancel={() =>
              this.setState({
                page:
                  this.state.selectedSolutionId !== null
                    ? 'admin'
                    : 'my-contributions',
                selectedSolutionId: null,
                isAdminViewingSolution: false
              })
            }
            onSubmit={this.handleAddSolutionSubmit}
          />
        </div>
      );
    }

    if (page === 'my-contributions') {
      return (
        <div className={styles.devFix}>
          {currentUser && (
            <DevFixNavbar
              currentUser={currentUser}
              onHome={this.handleHome}
              onSolutions={this.handleSolutions}
              onFavorites={this.handleFavorites}
              onMyContributions={this.handleMyContributions}
              onAddSolution={this.handleAddSolution}
              onAdmin={this.handleAdmin}
            />
          )}

          <MyContributions
            context={this.props.context}
            onBack={this.handleHome}
            onAddSolution={this.handleAddSolution}
            onViewSolution={(id: number) =>
              this.setState({
                page: 'details',
                selectedSolutionId: id,
                isAdminViewingSolution: false
              })
            }
          />
        </div>
      );
    }

    if (page === 'solutions') {
      return (
        <div className={styles.devFix}>
          {currentUser && (
            <DevFixNavbar
              currentUser={currentUser}
              onHome={this.handleHome}
              onSolutions={this.handleSolutions}
              onFavorites={this.handleFavorites}
              onMyContributions={this.handleMyContributions}
              onAddSolution={this.handleAddSolution}
              onAdmin={this.handleAdmin}
            />
          )}

          <Solutions
            context={this.props.context}
            onBack={this.handleHome}
            onViewSolution={(id: number) =>
              this.setState({
                page: 'details',
                selectedSolutionId: id,
                isAdminViewingSolution: false
              })
            }
          />
        </div>
      );
    }

    return (
      <div className={styles.devFix}>
        {currentUser && (
          <DevFixNavbar
            currentUser={currentUser}
            onHome={this.handleHome}
            onSolutions={this.handleSolutions}
            onFavorites={this.handleFavorites}
            onMyContributions={this.handleMyContributions}
            onAddSolution={this.handleAddSolution}
            onAdmin={this.handleAdmin}
          />
        )}

        <section className={styles.hero}>
          <div className={styles.heroContent}>
            <span className={styles.heroBadge}>
              <Sparkles size={14} />
              Developer Problem Solver
            </span>

            <h1>
              <span>Solve.</span> Learn. Share.
            </h1>

            <p>
              Find solutions to common development problems and
              share what you've learned with your team.
            </p>

            <div className={styles.searchContainer}>
              <div className={styles.searchInputWrapper}>
                <span className={styles.searchIcon}>🔍</span>

                <input
                  type="text"
                  placeholder="Search problems, errors, technologies..."
                  value={search}
                  onChange={event =>
                    this.setState({ search: event.target.value })
                  }
                />
              </div>

              <div className={styles.filterBar}>
                <select
                  value={technology}
                  onChange={event =>
                    this.setState({ technology: event.target.value })
                  }
                  aria-label="Technology"
                >
                  <option value="">All Technologies</option>
                  {technologies.map(item => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </select>

                <select
                  value={category}
                  onChange={event =>
                    this.setState({ category: event.target.value })
                  }
                  aria-label="Category"
                >
                  <option value="">All Categories</option>
                  {categories.map(item => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </select>

                <select
                  value={difficulty}
                  onChange={event =>
                    this.setState({ difficulty: event.target.value })
                  }
                  aria-label="Difficulty"
                >
                  <option value="">All Difficulties</option>
                  {difficulties.map(item => (
                    <option key={item} value={item}>
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
            </div>
          </div>
        </section>

        <main className={styles.mainContainer}>
          <section className={styles.statsGrid}>
            <div className={styles.statCard}>
              <BookOpen />
              <div>
                <span>Total Solutions</span>
                <strong>{published.length}</strong>
              </div>
            </div>

            <div className={styles.statCard}>
              <Layers />
              <div>
                <span>Technologies</span>
                <strong>{technologies.length}</strong>
              </div>
            </div>

            <div className={styles.statCard}>
              <ThumbsUp />
              <div>
                <span>Helpful Solutions</span>
                <strong>
                  {
                    published.filter(
                      solution => solution.HelpfulCount > 0
                    ).length
                  }
                </strong>
              </div>
            </div>

            <div className={styles.statCard}>
              <User />
              <div>
                <span>My Contributions</span>
                <strong>{myContributions}</strong>
              </div>
            </div>
          </section>

          {error && (
            <div className={styles.message}>{error}</div>
          )}

          {loading && (
            <div className={styles.message}>
              Loading solutions...
            </div>
          )}

          {!loading && !error && (
            <>
              <section className={styles.section}>
                <div className={styles.sectionHeader}>
                  <div>
                    <h2>
                      {search || technology || category || difficulty
                        ? `Search Results (${filtered.length})`
                        : 'Featured Solutions'}
                    </h2>

                    <p>
                      {search || technology || category || difficulty
                        ? 'Matching your search and filters'
                        : 'Hand-picked fixes the team keeps coming back to'}
                    </p>
                  </div>

                  <button
                    type="button"
                    onClick={this.handleSolutions}
                  >
                    Browse all →
                  </button>
                </div>

                {filtered.length > 0 ? (
                  <div className={styles.solutionGrid}>
                    {(search || technology || category || difficulty
                      ? filtered
                      : filtered.slice(0, 6)
                    ).map(solution => (
                      <div
                        key={solution.Id}
                        className={styles.solutionCard}
                      >
                        <div className={styles.cardTop}>
                          <div className={styles.cardBadges}>
                            <span className={styles.techBadge}>
                              {solution.Technology}
                            </span>

                            <span className={styles.categoryBadge}>
                              {solution.Category}
                            </span>

                            <span
                              className={this.getDifficultyClass(
                                solution.Difficulty
                              )}
                            >
                              {solution.Difficulty}
                            </span>
                          </div>

                          <button
                            type="button"
                            className={styles.favoriteButton}
                            aria-label="Add to favorites"
                            onClick={event =>
                              event.stopPropagation()
                            }
                          >
                            ⭐
                          </button>
                        </div>

                        <h3>{solution.Title}</h3>

                        <p>{solution.Problem}</p>

                        <div className={styles.tags}>
                          {(solution.Tags || '')
                            .split(',')
                            .map(tag => tag.trim())
                            .filter(Boolean)
                            .slice(0, 3)
                            .map(tag => (
                              <span key={tag} className={styles.tag}>
                                #{tag}
                              </span>
                            ))}
                        </div>

                        <div className={styles.cardActions}>
                          <button
                            type="button"
                            className={styles.helpfulButton}
                            onClick={event =>
                              event.stopPropagation()
                            }
                          >
                            <ThumbsUp size={14} />
                            {solution.HelpfulCount}
                          </button>

                          <button
                            type="button"
                            className={styles.viewSolutionButton}
                            onClick={event => {
                              event.stopPropagation();

                              this.setState({
                                page: 'details',
                                selectedSolutionId: solution.Id,
                                isAdminViewingSolution: false
                              });
                            }}
                          >
                            View Solution
                            <ArrowRight size={15} />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className={styles.emptyState}>
                    No solutions found.
                  </div>
                )}
              </section>

              <section className={styles.section}>
                <div className={styles.sectionHeader}>
                  <div>
                    <h2>Popular Technologies</h2>
                    <p>Jump straight into a stack</p>
                  </div>
                </div>

                <div className={styles.techGrid}>
                  {technologies.map(tech => (
                    <button
                      key={tech}
                      type="button"
                      className={styles.technologyCard}
                      onClick={() =>
                        this.setState({ technology: tech })
                      }
                    >
                      <div className={styles.techIcon}>
                        <Layers size={20} />
                      </div>

                      <strong>{tech}</strong>

                      <span>
                        {this.getTechnologyCount(tech)} solution
                        {this.getTechnologyCount(tech) === 1 ? '' : 's'}
                      </span>
                    </button>
                  ))}
                </div>
              </section>

              <section className={styles.bottomGrid}>
                <div>
                  <div className={styles.sectionHeader}>
                    <div>
                      <h2>Most Helpful Solutions</h2>
                      <p>Ranked by team votes</p>
                    </div>
                  </div>

                  <div className={styles.listCard}>
                    {mostHelpful.map((solution, index) => (
                      <div
                        key={solution.Id}
                        className={styles.listItem}
                      >
                        <strong>{index + 1}</strong>

                        <div>
                          <h4>{solution.Title}</h4>
                          <span>
                            {solution.Technology}
                            {' · '}
                            {solution.Difficulty}
                          </span>
                        </div>

                        <b>👍 {solution.HelpfulCount}</b>
                      </div>
                    ))}
                  </div>
                </div>

                <div>
                  <div className={styles.sectionHeader}>
                    <div>
                      <h2>Recently Added</h2>
                      <p>Fresh from the team</p>
                    </div>
                  </div>

                  <div className={styles.listCard}>
                    {recent.map(solution => (
                      <div
                        key={solution.Id}
                        className={styles.listItem}
                      >
                        <div>
                          <h4>{solution.Title}</h4>
                          <span>
                            {solution.Technology}
                            {' · '}
                            {this.formatDate(solution.Created)}
                          </span>
                        </div>

                        <button
                          type="button"
                          className={styles.recentViewButton}
                          onClick={() =>
                            this.setState({
                              page: 'details',
                              selectedSolutionId: solution.Id,
                              isAdminViewingSolution: false
                            })
                          }
                        >
                          View
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              </section>
            </>
          )}
        </main>
      </div>
    );
  }
}