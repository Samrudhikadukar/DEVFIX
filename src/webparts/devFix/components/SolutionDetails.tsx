import * as React from 'react';
import styles from './SolutionDetails.module.scss';
import { UserService, ICurrentUser } from '../services/UserService';

import {
    AlertTriangle,
    ArrowLeft,
    Bug,
    CheckCircle2,
    Lightbulb,
    Search,
    Star,
    ThumbsUp
} from 'lucide-react';

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from '../services/SharePointService';

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

export default class SolutionDetails
    extends React.Component<
        ISolutionDetailsProps,
        ISolutionDetailsState
    > {

    private sharePointService: SharePointService;
    private userService: UserService;

    public state: ISolutionDetailsState = {
        solution: null,
        loading: true,
        error: '',
        helpfulLoading: false,
        currentUser: null,
        favoriteLoading: false,
    };

    constructor(props: ISolutionDetailsProps) {
        super(props);

        this.sharePointService =
            new SharePointService(props.context);

        this.userService =
            new UserService(props.context);
    }

    public async componentDidMount(): Promise<void> {
        try {
            const currentUser =
                await this.userService.getCurrentUser();

            this.setState(
                {
                    currentUser
                },
                async () => {
                    await this.loadSolution();
                }
            );
        } catch (error) {
            console.error(
                'Error loading current user:',
                error
            );

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
                isFavorite =
                    await this.sharePointService.isFavorite(
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
            console.error(
                'Error loading solution details:',
                error
            );

            this.setState({
                loading: false,
                error: 'Unable to load solution.'
            });
        }
    }

    private async handleHelpful(): Promise<void> {
        if (
            !this.state.solution ||
            this.state.helpfulLoading
        ) {
            return;
        }

        try {
            this.setState({
                helpfulLoading: true
            });

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
            console.error(
                'Error marking solution as helpful:',
                error
            );

            this.setState({
                helpfulLoading: false
            });
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
            this.setState({
                favoriteLoading: true
            });

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
            console.error(
                'Error updating favorite:',
                error
            );

            this.setState({
                favoriteLoading: false
            });
        }
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

    private formatDate(date: string): string {
        if (!date) {
            return '';
        }

        return new Date(date).toLocaleDateString(
            'en-US',
            {
                month: 'short',
                day: 'numeric',
                year: 'numeric'
            }
        );
    }

    private renderBadge(
        text: string,
        className: string
    ): React.ReactElement {
        return (
            <span className={className}>
                {text}
            </span>
        );
    }

    private renderBlock(
        icon: React.ReactElement,
        title: string,
        content: string,
        toneClass: string,
        isCode: boolean = false
    ): React.ReactElement {
        return (
            <section className={styles.block}>
                <h2 className={styles.blockTitle}>
                    <span
                        className={`${styles.blockIcon} ${toneClass}`}
                    >
                        {icon}
                    </span>

                    {title}
                </h2>

                {isCode ? (
                    <pre className={styles.errorMessage}>
                        {content}
                    </pre>
                ) : (
                    <p className={styles.blockContent}>
                        {content}
                    </p>
                )}
            </section>
        );
    }

    public render(): React.ReactElement<ISolutionDetailsProps> {
        const {
            solution,
            loading,
            error
        } = this.state;

        if (loading) {
            return (
                <section className={styles.detailsPage}>
                    <div className={styles.message}>
                        Loading solution...
                    </div>
                </section>
            );
        }

        if (error || !solution) {
            return (
                <section className={styles.detailsPage}>
                    <div className={styles.message}>
                        <h2>Solution not found</h2>

                        <p>
                            This solution may have been deleted.
                        </p>

                        <button
                            type="button"
                            className={styles.backButton}
                            onClick={this.props.onBack}
                        >
                            <ArrowLeft size={16} />
                            Back to solutions
                        </button>
                    </div>
                </section>
            );
        }

        const tags = (solution.Tags || '')
            .split(',')
            .map(tag => tag.trim())
            .filter(Boolean);

        return (
            <section className={styles.detailsPage}>

                <div className={styles.content}>

                    {/* BACK */}
                    <button
                        type="button"
                        className={styles.backLink}
                        onClick={this.props.onBack}
                    >
                        <ArrowLeft size={16} />
                        All solutions
                    </button>

                    {/* HEADER */}
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
                                this.getDifficultyClass(
                                    solution.Difficulty
                                )
                            )}

                            {solution.Status !== 'Published' &&
                                this.renderBadge(
                                    solution.Status,
                                    styles.statusBadge
                                )}

                        </div>

                        <h1>
                            {solution.Title}
                        </h1>

                        <p className={styles.author}>
                            by{' '}
                            <strong>
                                {solution.Author?.Title || 'Unknown'}
                            </strong>
                            {' · '}
                            {this.formatDate(solution.Created)}
                        </p>

                        {/* ACTIONS */}
                        <div className={styles.actions}>

                            {/* HELPFUL */}
                            <button
                                type="button"
                                className={styles.helpfulButton}
                                onClick={() => this.handleHelpful()}
                                disabled={this.state.helpfulLoading}
                            >
                                <ThumbsUp size={16} />

                                Helpful ·{' '}
                                {solution.HelpfulCount || 0}
                            </button>

                            {/* FAVORITE */}
                            <button
                                type="button"
                                className={styles.favoriteButton}
                                onClick={() => this.handleFavorite()}
                                disabled={this.state.favoriteLoading}
                            >
                                <Star size={16} />

                                {solution.IsFavorite
                                    ? 'Saved to Favorites'
                                    : 'Add to Favorites'}
                            </button>

                        </div>

                    </header>

                    {/* DETAILS */}
                    <div className={styles.blocks}>

                        {/* PROBLEM */}
                        {this.renderBlock(
                            <Bug size={16} />,
                            'Problem',
                            solution.Problem,
                            styles.dangerTone
                        )}

                        {/* ERROR */}
                        {solution.ErrorMessage && (
                            this.renderBlock(
                                <AlertTriangle size={16} />,
                                'Error Message',
                                solution.ErrorMessage,
                                styles.warningTone,
                                true
                            )
                        )}

                        {/* CAUSE */}
                        {this.renderBlock(
                            <Search size={16} />,
                            'Cause',
                            solution.Cause,
                            styles.violetTone
                        )}

                        {/* SOLUTION */}
                        {this.renderBlock(
                            <CheckCircle2 size={16} />,
                            'Solution',
                            solution.Solution,
                            styles.successTone
                        )}

                        {/* LESSON */}
                        {solution.LessonLearned && (
                            this.renderBlock(
                                <Lightbulb size={16} />,
                                'Lesson Learned',
                                solution.LessonLearned,
                                styles.primaryTone
                            )
                        )}

                        {/* TAGS */}
                        {tags.length > 0 && (
                            <div className={styles.tags}>
                                {tags.map(tag => (
                                    <span
                                        key={tag}
                                        className={styles.tag}
                                    >
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