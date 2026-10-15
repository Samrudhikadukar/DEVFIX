
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
  activeTab: string;
}

export default class DevFixNavbar extends React.Component<
  IDevFixNavbarProps,
  IDevFixNavbarState
> {
  public state: IDevFixNavbarState = {
    isMobileMenuOpen: false,
    activeTab: 'Home'
  };

  private handleNavigation = (
    event: React.MouseEvent<HTMLElement>,
    callback: () => void,
    tabName: string
  ): void => {
    event.preventDefault();
    callback();

    this.setState({
      isMobileMenuOpen: false,
      activeTab: tabName
    });
  };

  private toggleMobileMenu = (): void => {
    this.setState(previousState => ({
      isMobileMenuOpen: !previousState.isMobileMenuOpen
    }));
  };

  public render(): React.ReactElement<IDevFixNavbarProps> {
    const { currentUser } = this.props;
    const { isMobileMenuOpen, activeTab } = this.state;

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
            className={`${styles.navLink} ${
              activeTab === 'Home' ? styles.activeNavLink : ''
            }`}
            onClick={event =>
              this.handleNavigation(event, this.props.onHome, 'Home')
            }
          >
            <Icon iconName="Home" className={styles.navIcon} />
            Home
          </Link>

          <Link
            href="#"
            className={`${styles.navLink} ${
              activeTab === 'Solutions' ? styles.activeNavLink : ''
            }`}
            onClick={event =>
              this.handleNavigation(
                event,
                this.props.onSolutions,
                'Solutions'
              )
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
            className={`${styles.navLink} ${
              activeTab === 'Favorites' ? styles.activeNavLink : ''
            }`}
            onClick={event =>
              this.handleNavigation(
                event,
                this.props.onFavorites,
                'Favorites'
              )
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
            className={`${styles.navLink} ${
              activeTab === 'My Contributions'
                ? styles.activeNavLink
                : ''
            }`}
            onClick={event =>
              this.handleNavigation(
                event,
                this.props.onMyContributions,
                'My Contributions'
              )
            }
          >
            <Icon iconName="Contact" className={styles.navIcon} />
            My Contributions
          </Link>

          {currentUser.role === 'ADMIN' && (
            <Link
              href="#"
              className={`${styles.navLink} ${
                activeTab === 'Admin' ? styles.activeNavLink : ''
              }`}
              onClick={event =>
                this.handleNavigation(
                  event,
                  this.props.onAdmin,
                  'Admin'
                )
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
