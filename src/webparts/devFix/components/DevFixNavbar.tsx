
import * as React from 'react';
import { Link, PrimaryButton, Text } from '@fluentui/react';
import { ICurrentUser } from '../services/UserService';
import styles from './DevFixNavbar.module.scss';

interface IDevFixNavbarProps {
  currentUser: ICurrentUser;
  onHome: () => void;
  onSolutions: () => void;
  onFavorites: () => void;
  onMyContributions: () => void;
  onAddSolution: () => void;
  onAdmin: () => void;
}

export default class DevFixNavbar extends React.Component<IDevFixNavbarProps> {
  public render(): React.ReactElement<IDevFixNavbarProps> {
    const { currentUser } = this.props;

    const handleNavigation = (
      event: React.MouseEvent<HTMLAnchorElement | HTMLElement | HTMLButtonElement>,
      callback: () => void
    ): void => {
      event.preventDefault();
      callback();
    };

    return (
      <header className={styles.navbar}>
        <div className={styles.logoSection}>
          <div className={styles.logoIcon}>
            🔧
          </div>

          <div className={styles.logoText}>
            <Text className={styles.logoTitle} variant="large">
              DevFix
            </Text>

            <Text className={styles.logoSubtitle}>
              Developer Problem Solver
            </Text>
          </div>
        </div>

        <nav className={styles.navigation}>
          <Link
            href="#"
            className={styles.navLink}
            onClick={(event) => handleNavigation(event, this.props.onHome)}
          >
            Home
          </Link>

          <Link
            href="#"
            className={styles.navLink}
            onClick={(event) => handleNavigation(event, this.props.onSolutions)}
          >
            Solutions
          </Link>

          <Link
            href="#"
            className={styles.navLink}
            onClick={(event) => handleNavigation(event, this.props.onFavorites)}
          >
            Favorites
          </Link>

          <Link
            href="#"
            className={styles.navLink}
            onClick={(event) => handleNavigation(event, this.props.onMyContributions)}
          >
            My Contributions
          </Link>

          {currentUser.role === 'ADMIN' && (
            <Link
              href="#"
              className={styles.navLink}
              onClick={(event) => handleNavigation(event, this.props.onAdmin)}
            >
              Admin
            </Link>
          )}
        </nav>

        <div className={styles.rightSection}>
          <PrimaryButton
            className={styles.addButton}
            text="+ Add Solution"
            onClick={this.props.onAddSolution}
          />

          <div className={styles.userProfile}>
            <div className={styles.userInfo}>
              <Text className={styles.userName}>
                {currentUser.name}
              </Text>

              <Text
                className={`${styles.userRole} ${
                  currentUser.role === 'ADMIN'
                    ? styles.adminRole
                    : styles.developerRole
                }`}
              >
                {currentUser.role === 'ADMIN' ? 'Admin' : 'Developer'}
              </Text>
            </div>
          </div>
        </div>
      </header>
    );
  }
}
