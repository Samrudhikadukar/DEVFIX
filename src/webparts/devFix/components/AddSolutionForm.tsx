import * as React from 'react';
import styles from './Solutions.module.scss';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SharePointService } from "../services/SharePointService";

/* =====================================================
   INTERFACES
===================================================== */

interface IAddSolutionFormProps {
  onCancel: () => void;
  onSubmit: (
    solution: IAddSolutionInput
  ) => Promise<void>;
  initialSolution?: IAddSolutionInput;
  isEditMode?: boolean;
  context: WebPartContext;
}

export interface IAddSolutionInput {
  title: string;
  technology: string;
  category: string;
  difficulty: string;
  problem: string;
  errorMessage: string;
  cause: string;
  solution: string;
  lessonLearned: string;
  tags: string;
}

interface IFormState {
  title: string;
  technology: string;
  category: string;
  difficulty: string;
  problem: string;
  errorMessage: string;
  cause: string;
  solution: string;
  lessonLearned: string;
  tags: string;
}

type FormErrors = Partial<
  Record<keyof IFormState, string>
>;

/* =====================================================
   OPTIONS
===================================================== */

const TECHNOLOGIES = [
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

const CATEGORIES = [
  'Debugging',
  'Deployment',
  'API',
  'Authentication',
  'Configuration',
  'Permissions',
  'UI/UX',
  'Performance'
];

const DIFFICULTIES = [
  'Easy',
  'Medium',
  'Hard'
];

const REQUIRED: (keyof IFormState)[] = [
  'title',
  'technology',
  'category',
  'difficulty',
  'problem',
  'cause',
  'solution'
];

/* =====================================================
   COMPONENT
===================================================== */

export default class AddSolutionForm extends React.Component<
  IAddSolutionFormProps,
  {
    form: IFormState;
    errors: FormErrors;
    saving: boolean;
     technologyChoices: string[];
    categoryChoices: string[];
    difficultyChoices: string[];
  }
> {

  /* =====================================================
     EMPTY FORM
  ===================================================== */

  private getEmptyForm = (): IFormState => ({
    title: '',
    technology: '',
    category: '',
    difficulty: '',
    problem: '',
    errorMessage: '',
    cause: '',
    solution: '',
    lessonLearned: '',
    tags: ''
  });

  /* =====================================================
     STATE
  ===================================================== */

  public state = {
    form: this.getEmptyForm(),
    errors: {} as FormErrors,
    saving: false,
     technologyChoices: [],
  categoryChoices: [],
  difficultyChoices: []
  };

  /* =====================================================
     LOAD EDIT DATA
  ===================================================== */

  public async componentDidMount(): Promise<void> {

  if (this.props.initialSolution) {

    this.setState({

      form: {

        ...this.props.initialSolution

      }

    });

  }

  const service = new SharePointService(
    this.props.context
  );

  const [
    technologies,
    categories,
    difficulties
  ] = await Promise.all([
    service.getChoiceValues("Technology"),
    service.getChoiceValues("Category"),
    service.getChoiceValues("Difficulty")
  ]);

  this.setState({
    technologyChoices: technologies,
    categoryChoices: categories,
    difficultyChoices: difficulties
  });
}

  public componentDidUpdate(
    prevProps: IAddSolutionFormProps
  ): void {

    if (
      this.props.initialSolution &&
      this.props.initialSolution !==
        prevProps.initialSolution
    ) {
      this.setState({
        form: {
          ...this.props.initialSolution
        },
        errors: {}
      });
    }
  }

  /* =====================================================
     HANDLE CHANGE
  ===================================================== */

  private handleChange =
    (field: keyof IFormState) =>
    (
      event: React.ChangeEvent<
        HTMLInputElement |
        HTMLTextAreaElement |
        HTMLSelectElement
      >
    ): void => {

      const value =
        event.target.value;

      this.setState(prevState => ({
        form: {
          ...prevState.form,
          [field]: value
        },

        errors: {
          ...prevState.errors,
          [field]: undefined
        }
      }));
    };

  /* =====================================================
     VALIDATION
  ===================================================== */

  private validate(): boolean {

    const errors: FormErrors = {};

    REQUIRED.forEach(field => {

      if (
        !this.state.form[field]
          .trim()
      ) {
        errors[field] =
          'This field is required';
      }

    });

    if (
      this.state.form.title
        .trim() &&
      this.state.form.title
        .trim()
        .length < 10
    ) {
      errors.title =
        'Title should be at least 10 characters';
    }

    this.setState({
      errors
    });

    return (
      Object.keys(errors).length === 0
    );
  }

  /* =====================================================
     SUBMIT
  ===================================================== */

  private handleSubmit = async (
    event: React.FormEvent
  ): Promise<void> => {

    event.preventDefault();

    if (!this.validate()) {
      return;
    }

    try {

      this.setState({
        saving: true
      });

      await this.props.onSubmit({

        title:
          this.state.form.title.trim(),

        technology:
          this.state.form.technology,

        category:
          this.state.form.category,

        difficulty:
          this.state.form.difficulty,

        problem:
          this.state.form.problem.trim(),

        errorMessage:
          this.state.form.errorMessage.trim(),

        cause:
          this.state.form.cause.trim(),

        solution:
          this.state.form.solution.trim(),

        lessonLearned:
          this.state.form.lessonLearned.trim(),

        tags:
          this.state.form.tags.trim()
      });

    } catch (error) {

      console.error(
        'Error saving solution:',
        error
      );

      this.setState({
        saving: false
      });
    }
  };

  /* =====================================================
     FIELD ERROR
  ===================================================== */

  private renderFieldError(
    field: keyof IFormState
  ): React.ReactElement | null {

    const error =
      this.state.errors[field];

    if (!error) {
      return null;
    }

    return (
      <span
        className={styles.formError}
      >
        {error}
      </span>
    );
  }

  /* =====================================================
     LABEL
  ===================================================== */

  private renderLabel(
    field: keyof IFormState,
    label: string
  ): React.ReactElement {

    return (
      <span
        className={styles.formLabel}
      >

        {label}

        {REQUIRED.indexOf(field) !== -1 && (
          <span
            className={styles.required}
          >
            *
          </span>
        )}

      </span>
    );
  }

  /* =====================================================
     RENDER
  ===================================================== */

  public render(): React.ReactElement {

    const {
      form,
      saving
    } = this.state;

    const isEditMode =
      this.props.isEditMode === true;

    return (
      <div
        className={
          styles.addSolutionPage
        }
      >

        {/* HEADER */}

        <div
          className={
            styles.addSolutionHeader
          }
        >

          <div>

            <h1>
              {isEditMode
                ? 'Edit Solution'
                : 'Add Solution'}
            </h1>

            <p>
              {isEditMode
                ? 'Update the solution details.'
                : 'Share your solution with the team.'}
            </p>

          </div>

        </div>

        {/* FORM */}

        <form
          onSubmit={this.handleSubmit}
          className={
            styles.addSolutionForm
          }
          noValidate
        >

          {/* TITLE */}

          <div
            className={
              styles.formField
            }
          >

            {this.renderLabel(
              'title',
              'Solution Title'
            )}

            <input
              className={`
                ${styles.formInput}
                ${
                  this.state.errors.title
                    ? styles.formInputError
                    : ''
                }
              `}
              value={form.title}
              onChange={
                this.handleChange('title')
              }
              placeholder="e.g. PnPjs data not loading from list"
            />

            {this.renderFieldError(
              'title'
            )}

          </div>

          {/* TECHNOLOGY / CATEGORY / DIFFICULTY */}

          <div
            className={
              styles.formThreeColumns
            }
          >

            {/* TECHNOLOGY */}

            <div
              className={
                styles.formField
              }
            >

              {this.renderLabel(
                'technology',
                'Technology'
              )}

              <select
                className={
                  styles.formInput
                }
                value={form.technology}
                onChange={
                  this.handleChange(
                    'technology'
                  )
                }
              >

                <option value="">
                  Select…
                </option>

                {this.state.technologyChoices.map(item => (
  <option key={item} value={item}>
    {item}
  </option>
))}

              </select>

              {this.renderFieldError(
                'technology'
              )}

            </div>

            {/* CATEGORY */}

            <div
              className={
                styles.formField
              }
            >

              {this.renderLabel(
                'category',
                'Category'
              )}

              <select
                className={
                  styles.formInput
                }
                value={form.category}
                onChange={
                  this.handleChange(
                    'category'
                  )
                }
              >

                <option value="">
                  Select…
                </option>

               {this.state.categoryChoices.map(item => (
  <option key={item} value={item}>
    {item}
  </option>
))}

              </select>

              {this.renderFieldError(
                'category'
              )}

            </div>

            {/* DIFFICULTY */}

            <div
              className={
                styles.formField
              }
            >

              {this.renderLabel(
                'difficulty',
                'Difficulty'
              )}

              <select
                className={
                  styles.formInput
                }
                value={form.difficulty}
                onChange={
                  this.handleChange(
                    'difficulty'
                  )
                }
              >

                <option value="">
                  Select…
                </option>

                {this.state.difficultyChoices.map(item => (
  <option key={item} value={item}>
    {item}
  </option>
))}

              </select>

              {this.renderFieldError(
                'difficulty'
              )}

            </div>

          </div>

          {/* PROBLEM */}

          <div
            className={
              styles.formField
            }
          >

            {this.renderLabel(
              'problem',
              'Problem'
            )}

            <textarea
              rows={3}
              className={
                styles.formTextarea
              }
              value={form.problem}
              onChange={
                this.handleChange(
                  'problem'
                )
              }
              placeholder="What went wrong?"
            />

            {this.renderFieldError(
              'problem'
            )}

          </div>

          {/* ERROR MESSAGE */}

          <div
            className={
              styles.formField
            }
          >

            {this.renderLabel(
              'errorMessage',
              'Error Message'
            )}

            <textarea
              rows={2}
              className={`
                ${styles.formTextarea}
                ${styles.codeInput}
              `}
              value={
                form.errorMessage
              }
              onChange={
                this.handleChange(
                  'errorMessage'
                )
              }
              placeholder="Paste the exact error, if any"
            />

            {this.renderFieldError(
              'errorMessage'
            )}

          </div>

          {/* CAUSE */}

          <div
            className={
              styles.formField
            }
          >

            {this.renderLabel(
              'cause',
              'Cause'
            )}

            <textarea
              rows={2}
              className={
                styles.formTextarea
              }
              value={form.cause}
              onChange={
                this.handleChange('cause')
              }
              placeholder="Why did it happen?"
            />

            {this.renderFieldError(
              'cause'
            )}

          </div>

          {/* SOLUTION */}

          <div
            className={
              styles.formField
            }
          >

            {this.renderLabel(
              'solution',
              'Solution'
            )}

            <textarea
              rows={4}
              className={
                styles.formTextarea
              }
              value={form.solution}
              onChange={
                this.handleChange(
                  'solution'
                )
              }
              placeholder="How did you fix it? Steps, code, settings…"
            />

            {this.renderFieldError(
              'solution'
            )}

          </div>

          {/* LESSON */}

          <div
            className={
              styles.formField
            }
          >

            {this.renderLabel(
              'lessonLearned',
              'Lesson Learned'
            )}

            <textarea
              rows={2}
              className={
                styles.formTextarea
              }
              value={
                form.lessonLearned
              }
              onChange={
                this.handleChange(
                  'lessonLearned'
                )
              }
              placeholder="What should others remember?"
            />

            {this.renderFieldError(
              'lessonLearned'
            )}

          </div>

          {/* TAGS */}

          <div
            className={
              styles.formField
            }
          >

            {this.renderLabel(
              'tags',
              'Tags'
            )}

            <input
              className={
                styles.formInput
              }
              value={form.tags}
              onChange={
                this.handleChange('tags')
              }
              placeholder="spfx, deployment, gulp"
            />

            <span
              className={
                styles.formHint
              }
            >
              Comma-separated, e.g. spfx,
              deployment, gulp
            </span>

          </div>

          {/* ACTIONS */}

          <div
            className={
              styles.formActions
            }
          >

            <button
              type="button"
              className={
                styles.cancelButton
              }
              onClick={
                this.props.onCancel
              }
              disabled={saving}
            >
              Cancel
            </button>

            <button
              type="submit"
              className={
                styles.saveButton
              }
              disabled={saving}
            >
              {saving
                ? 'Saving...'
                : isEditMode
                  ? 'Update Solution'
                  : 'Save Solution'}
            </button>

          </div>

        </form>

      </div>
    );
  }
}