
import * as React from 'react';
import {
  Icon,
  Text,
  Link,
  PrimaryButton,
  IconButton
} from '@fluentui/react';

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

interface IDevFixNavbarState {
  isMobileMenuOpen: boolean;
}

export default class DevFixNavbar extends React.Component<
  IDevFixNavbarProps,
  IDevFixNavbarState
> {
  public state: IDevFixNavbarState = {
    isMobileMenuOpen: false
  };

  private handleNavigation = (
    event: React.MouseEvent<HTMLElement>,
    callback: () => void
  ): void => {
    event.preventDefault();
    callback();

    this.setState({
      isMobileMenuOpen: false
    });
  };

  private toggleMobileMenu = (): void => {
    this.setState(previousState => ({
      isMobileMenuOpen: !previousState.isMobileMenuOpen
    }));
  };

  public render(): React.ReactElement<IDevFixNavbarProps> {
    const { currentUser } = this.props;
    const { isMobileMenuOpen } = this.state;

    return (
      <header className={styles.navbar}>
        {/* Logo */}
        <div className={styles.logoSection}>
          <div className={styles.logoIcon}>
            <Icon iconName="Build" />
          </div>

          <div className={styles.logoText}>
            <Text className={styles.logoTitle}>
              DevFix
            </Text>

            <Text className={styles.logoSubtitle}>
              Developer Problem Solver
            </Text>
          </div>
        </div>

        {/* Mobile hamburger button */}
        <IconButton
          className={styles.menuToggle}
          iconProps={{
            iconName: isMobileMenuOpen ? 'Cancel' : 'GlobalNavButton'
          }}
          ariaLabel={
            isMobileMenuOpen
              ? 'Close navigation menu'
              : 'Open navigation menu'
          }
          aria-expanded={isMobileMenuOpen}
          onClick={this.toggleMobileMenu}
        />

        {/* Navigation */}
        <nav
          className={`${styles.navigation} ${
            isMobileMenuOpen ? styles.navigationOpen : ''
          }`}
          aria-label="Main navigation"
        >
          <Link
            href="#"
            className={styles.navLink}
            onClick={event =>
              this.handleNavigation(event, this.props.onHome)
            }
          >
            <Icon iconName="Home" className={styles.navIcon} />
            Home
          </Link>

          <Link
            href="#"
            className={styles.navLink}
            onClick={event =>
              this.handleNavigation(event, this.props.onSolutions)
            }
          >
            <Icon
              iconName="KnowledgeArticle"
              className={styles.navIcon}
            />
            Solutions
          </Link>

          <Link
            href="#"
            className={styles.navLink}
            onClick={event =>
              this.handleNavigation(event, this.props.onFavorites)
            }
          >
            <Icon
              iconName="FavoriteStar"
              className={styles.navIcon}
            />
            Favorites
          </Link>

          <Link
            href="#"
            className={styles.navLink}
            onClick={event =>
              this.handleNavigation(
                event,
                this.props.onMyContributions
              )
            }
          >
            <Icon iconName="Contact" className={styles.navIcon} />
            My Contributions
          </Link>

          {currentUser.role === 'ADMIN' && (
            <Link
              href="#"
              className={styles.navLink}
              onClick={event =>
                this.handleNavigation(event, this.props.onAdmin)
              }
            >
              <Icon iconName="Shield" className={styles.navIcon} />
              Admin
            </Link>
          )}
        </nav>

        {/* Add Solution and user profile */}
        <div className={styles.rightSection}>
          <PrimaryButton
            className={styles.addButton}
            text="Add Solution"
            iconProps={{ iconName: 'Add' }}
            onClick={() => {
              this.setState({ isMobileMenuOpen: false });
              this.props.onAddSolution();
            }}
          />

          <div className={styles.userProfile}>
            <div className={styles.avatar}>
              {currentUser.initials}
            </div>

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
                {currentUser.role === 'ADMIN'
                  ? 'Admin'
                  : 'Developer'}
              </Text>
            </div>
          </div>
        </div>
      </header>
    );
  }
}