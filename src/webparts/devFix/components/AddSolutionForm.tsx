
import * as React from 'react';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton
} from '@fluentui/react';

import styles from './Solutions.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from '../services/SharePointService';

/* =====================================================
   INTERFACES
===================================================== */

interface IAddSolutionFormProps {
  onCancel: () => void;
  onSubmit: (solution: IAddSolutionInput) => Promise<void>;
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

type FormErrors = Partial<Record<keyof IFormState, string>>;

/* =====================================================
   REQUIRED FIELDS
===================================================== */

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
    technologyChoices: [] as string[],
    categoryChoices: [] as string[],
    difficultyChoices: [] as string[]
  };

  /* =====================================================
     LOAD EDIT DATA AND SHAREPOINT CHOICES
  ===================================================== */

  public async componentDidMount(): Promise<void> {
    if (this.props.initialSolution) {
      this.setState({
        form: { ...this.props.initialSolution }
      });
    }

    const service = new SharePointService(this.props.context);

    const [technologies, categories, difficulties] = await Promise.all([
      service.getChoiceValues('Technology'),
      service.getChoiceValues('Category'),
      service.getChoiceValues('Difficulty')
    ]);

    this.setState({
      technologyChoices: technologies,
      categoryChoices: categories,
      difficultyChoices: difficulties
    });
  }

  public componentDidUpdate(prevProps: IAddSolutionFormProps): void {
    if (
      this.props.initialSolution &&
      this.props.initialSolution !== prevProps.initialSolution
    ) {
      this.setState({
        form: { ...this.props.initialSolution },
        errors: {}
      });
    }
  }

  /* =====================================================
     FLUENT UI DROPDOWN OPTIONS
  ===================================================== */

  private getOptions = (values: string[]): IDropdownOption[] =>
    values.map(item => ({
      key: item,
      text: item
    }));

  /* =====================================================
     HANDLE FLUENT UI FIELD CHANGES
  ===================================================== */

  private handleChange = (
    field: keyof IFormState,
    value: string
  ): void => {
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
      if (!this.state.form[field].trim()) {
        errors[field] = 'This field is required';
      }
    });

    if (
      this.state.form.title.trim() &&
      this.state.form.title.trim().length < 10
    ) {
      errors.title = 'Title should be at least 10 characters';
    }

    this.setState({ errors });

    return Object.keys(errors).length === 0;
  }

  /* =====================================================
     SUBMIT
  ===================================================== */

  private handleSubmit = async (
    event: React.FormEvent<HTMLFormElement>
  ): Promise<void> => {
    event.preventDefault();

    if (!this.validate()) {
      return;
    }

    try {
      this.setState({ saving: true });

      await this.props.onSubmit({
        title: this.state.form.title.trim(),
        technology: this.state.form.technology,
        category: this.state.form.category,
        difficulty: this.state.form.difficulty,
        problem: this.state.form.problem.trim(),
        errorMessage: this.state.form.errorMessage.trim(),
        cause: this.state.form.cause.trim(),
        solution: this.state.form.solution.trim(),
        lessonLearned: this.state.form.lessonLearned.trim(),
        tags: this.state.form.tags.trim()
      });
    } catch (error) {
      console.error('Error saving solution:', error);

      this.setState({
        saving: false
      });
    }
  };

  /* =====================================================
     LABEL
  ===================================================== */

  private renderLabel(
    field: keyof IFormState,
    label: string
  ): React.ReactElement {
    return (
      <span className={styles.formLabel}>
        {label}

        {REQUIRED.indexOf(field) !== -1 && (
          <span className={styles.required}>*</span>
        )}
      </span>
    );
  }

  /* =====================================================
     FLUENT UI TEXT FIELD
  ===================================================== */

  private renderTextField = (
    field: keyof IFormState,
    label: string,
    placeholder: string,
    multiline: boolean = false,
    rows: number = 1
  ): React.ReactElement => {
    return (
      <div className={styles.formField}>
        {this.renderLabel(field, label)}

        <TextField
          className={styles.formInput}
          value={this.state.form[field]}
          placeholder={placeholder}
          multiline={multiline}
          rows={multiline ? rows : undefined}
          resizable={false}
          errorMessage={this.state.errors[field]}
          onChange={(_, value) =>
            this.handleChange(field, value || '')
          }
          disabled={this.state.saving}
          styles={{
            root: { width: '100%' },
            fieldGroup: {
              minHeight: multiline ? undefined : 42,
              borderRadius: 8,
              borderColor: this.state.errors[field]
                ? '#d13438'
                : '#d1d5db'
            },
            field: {
              padding: '8px 12px',
              fontSize: 14
            },
            errorMessage: {
              marginTop: 5
            }
          }}
        />
      </div>
    );
  };

  /* =====================================================
     FLUENT UI DROPDOWN
  ===================================================== */

  private renderDropdown = (
    field: 'technology' | 'category' | 'difficulty',
    label: string,
    choices: string[]
  ): React.ReactElement => {
    return (
      <div className={styles.formField}>
        {this.renderLabel(field, label)}

        <Dropdown
          placeholder="Select..."
          options={this.getOptions(choices)}
          selectedKey={this.state.form[field] || undefined}
          onChange={(_, option) =>
            this.handleChange(
              field,
              option ? String(option.key) : ''
            )
          }
          errorMessage={this.state.errors[field]}
          disabled={this.state.saving}
          styles={{
            root: { width: '100%' },
            title: {
              height: 42,
              lineHeight: 40,
              borderRadius: 8,
              borderColor: this.state.errors[field]
                ? '#d13438'
                : '#d1d5db',
              paddingLeft: 12
            },
            dropdown: { width: '100%' },
            caretDownWrapper: { height: 40 },
            errorMessage: { marginTop: 5 }
          }}
        />
      </div>
    );
  };

  /* =====================================================
     RENDER
  ===================================================== */

  public render(): React.ReactElement {
    const { form, saving } = this.state;
    const isEditMode = this.props.isEditMode === true;

    return (
      <div className={styles.addSolutionPage}>
        {/* HEADER */}

        <div className={styles.addSolutionHeader}>
          <div>
            <h1>
              {isEditMode ? 'Edit Solution' : 'Add Solution'}
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
          className={styles.addSolutionForm}
          noValidate
        >
          {/* TITLE */}

          {this.renderTextField(
            'title',
            'Solution Title',
            'e.g. PnPjs data not loading from list'
          )}

          {/* TECHNOLOGY / CATEGORY / DIFFICULTY */}

          <div className={styles.formThreeColumns}>
            {this.renderDropdown(
              'technology',
              'Technology',
              this.state.technologyChoices
            )}

            {this.renderDropdown(
              'category',
              'Category',
              this.state.categoryChoices
            )}

            {this.renderDropdown(
              'difficulty',
              'Difficulty',
              this.state.difficultyChoices
            )}
          </div>

          {/* PROBLEM */}

          {this.renderTextField(
            'problem',
            'Problem',
            'What went wrong?',
            true,
            3
          )}

          {/* ERROR MESSAGE */}

          {this.renderTextField(
            'errorMessage',
            'Error Message',
            'Paste the exact error, if any',
            true,
            2
          )}

          {/* CAUSE */}

          {this.renderTextField(
            'cause',
            'Cause',
            'Why did it happen?',
            true,
            2
          )}

          {/* SOLUTION */}

          {this.renderTextField(
            'solution',
            'Solution',
            'How did you fix it? Steps, code, settings…',
            true,
            4
          )}

          {/* LESSON LEARNED */}

          {this.renderTextField(
            'lessonLearned',
            'Lesson Learned',
            'What should others remember?',
            true,
            2
          )}

          {/* TAGS */}

          <div className={styles.formField}>
            {this.renderLabel('tags', 'Tags')}

            <TextField
              className={styles.formInput}
              value={form.tags}
              onChange={(_, value) =>
                this.handleChange('tags', value || '')
              }
              placeholder="spfx, deployment, gulp"
              disabled={saving}
              styles={{
                root: { width: '100%' },
                fieldGroup: {
                  height: 42,
                  borderRadius: 8,
                  borderColor: '#d1d5db'
                },
                field: {
                  padding: '8px 12px',
                  fontSize: 14
                }
              }}
            />

            <span className={styles.formHint}>
              Comma-separated, e.g. spfx, deployment, gulp
            </span>
          </div>

          {/* ACTIONS */}

          <div className={styles.formActions}>
            <DefaultButton
              text="Cancel"
              onClick={this.props.onCancel}
              disabled={saving}
              className={styles.cancelButton}
              styles={{
                root: {
                  minWidth: 110,
                  height: 42,
                  borderRadius: 8
                }
              }}
            />

            <PrimaryButton
              type="submit"
              text={
                saving
                  ? 'Saving...'
                  : isEditMode
                    ? 'Update Solution'
                    : 'Save Solution'
              }
              disabled={saving}
              className={styles.saveButton}
              styles={{
                root: {
                  minWidth: 150,
                  height: 42,
                  borderRadius: 8,
                  backgroundColor: '#7c3aed',
                  borderColor: '#7c3aed'
                },
                rootHovered: {
                  backgroundColor: '#6d28d9',
                  borderColor: '#6d28d9'
                }
              }}
            />
          </div>
        </form>
      </div>
    );
  }
}

