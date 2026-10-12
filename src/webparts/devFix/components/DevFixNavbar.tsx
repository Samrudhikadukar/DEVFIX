import * as React from 'react';
import { ICurrentUser } from '../services/UserService';
import styles from './DevFixNavbar.module.scss';

interface IDevFixNavbarProps {
  currentUser: ICurrentUser;
  onHome: () => void;
  onSolutions: () => void;
  onMyContributions: () => void;
  onAddSolution: () => void;
  onAdmin: () => void;
}

export default class DevFixNavbar extends React.Component<IDevFixNavbarProps> {

  public render(): React.ReactElement<IDevFixNavbarProps> {
    const { currentUser } = this.props;

    return (
      <header className={styles.navbar}>

        <div className={styles.logoSection}>
          <div className={styles.logoIcon}>
            🔧
          </div>

          <div className={styles.logoText}>
            <div className={styles.logoTitle}>
              DevFix
            </div>

            <div className={styles.logoSubtitle}>
              Developer Problem Solver
            </div>
          </div>
        </div>

        <nav className={styles.navigation}>

          <a
            href="#"
            className={styles.navLink}
            onClick={(event) => {
              event.preventDefault();
              this.props.onHome();
            }}
          >
            Home
          </a>

          <a
            href="#"
            className={styles.navLink}
            onClick={(event) => {
              event.preventDefault();
              this.props.onSolutions();
            }}
          >
            Solutions
          </a>

          <a
            href="#"
            className={styles.navLink}
            onClick={(event) => {
              event.preventDefault();
            }}
          >
            Favorites
          </a>

          <a
            href="#"
            className={styles.navLink}
            onClick={(event) => {
              event.preventDefault();
              this.props.onMyContributions();
            }}
          >
            My Contributions
          </a>

          {currentUser.role === 'ADMIN' && (
            <a
              href="#"
              className={styles.navLink}
              onClick={(event) => {
                event.preventDefault();
                this.props.onAdmin();
              }}
            >
                               Admin
            </a>
          )}

        </nav>

        <div className={styles.rightSection}>

          <button
            type="button"
            className={styles.addButton}
            onClick={this.props.onAddSolution}
          >
            + Add Solution
          </button>

          <div className={styles.userProfile}>

            {/* <div className={styles.avatar}>
              {currentUser.initials}
            </div> */}

            <div className={styles.userInfo}>
              <div className={styles.userName}>
                {currentUser.name}
              </div>

              <div
                className={`${styles.userRole} ${
                  currentUser.role === 'ADMIN'
                    ? styles.adminRole
                    : styles.developerRole
                }`}
              >
                {currentUser.role === 'ADMIN'
                  ? 'Admin'
                  : 'Developer'}
              </div>
            </div>

          </div>

        </div>

      </header>
    );
  }
}