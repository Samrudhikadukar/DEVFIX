import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { UserService } from "../services/UserService";
import styles from "./ManageUsers.module.scss";

interface IUser {
  id: number;
  name: string;
  email: string;
  loginName: string;
  role: "USER" | "ADMIN";
}

interface ISolution {
  Id: number;
  Author?: {
    Id: number;
    Title: string;
    EMail: string;
  };
}

export interface IManageUsersProps {
  context: WebPartContext;
  solutions: ISolution[];
  onBack: () => void;
}

const ManageUsers: React.FC<IManageUsersProps> = ({
  context,
  solutions,
  onBack
}) => {
  const [users, setUsers] = useState<IUser[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>("");

  useEffect(() => {
    const loadUsers = async (): Promise<void> => {
      try {
        setLoading(true);
        setError("");

        const userService = new UserService(context);
        const data = await userService.getUsers();

        setUsers(data as IUser[]);
      } catch (err) {
        console.error("Error loading users:", err);
        setError("Unable to load users.");
      } finally {
        setLoading(false);
      }
    };

    loadUsers().catch((err) => {
      console.error(err);
    });
  }, [context]);

  const getSolutionCount = (userId: number): number => {
    return solutions.filter(
      (solution) => solution.Author?.Id === userId
    ).length;
  };

  const getRoleClass = (role: "USER" | "ADMIN"): string => {
    return role === "ADMIN"
      ? styles.adminBadge
      : styles.userBadge;
  };

  return (
    <div className={styles.manageUsers}>

      {/* Header */}
      <div className={styles.headerSection}>

        <div className={styles.headerContent}>
          <h1 className={styles.pageTitle}>
            Manage Users
          </h1>

          <p className={styles.pageSubtitle}>
            Members of the DevFix knowledge base and their roles
          </p>
        </div>

        <button
          type="button"
          className={styles.backButton}
          onClick={onBack}
        >
          ← Back to Admin Dashboard
        </button>

      </div>

      {/* Users Table */}
      <div className={styles.tableContainer}>

        {/* Loading */}
        {loading && (
          <div className={styles.message}>
            Loading users...
          </div>
        )}

        {/* Error */}
        {!loading && error && (
          <div className={styles.errorMessage}>
            {error}
          </div>
        )}

        {/* Table */}
        {!loading && !error && (
          <table className={styles.usersTable}>

            <thead>
              <tr>
                <th>User</th>
                <th>Email</th>
                <th>Role</th>
                <th>Solutions</th>
              </tr>
            </thead>

            <tbody>

              {users.map((user) => (
                <tr key={user.id}>

                  {/* User */}
                  <td>
                    <span className={styles.userName}>
                      {user.name || "Unknown User"}
                    </span>
                  </td>

                  {/* Email */}
                  <td>
                    <span className={styles.email}>
                      {user.email || "-"}
                    </span>
                  </td>

                  {/* Role */}
                  <td>
                    <span
                      className={`${styles.roleBadge} ${getRoleClass(
                        user.role
                      )}`}
                    >
                      {user.role === "ADMIN" ? "Admin" : "User"}
                    </span>
                  </td>

                  {/* Solutions */}
                  <td>
                    <span className={styles.solutionCount}>
                      {getSolutionCount(user.id)}
                    </span>
                  </td>

                </tr>
              ))}

              {/* No users */}
              {users.length === 0 && (
                <tr>
                  <td
                    colSpan={4}
                    className={styles.emptyMessage}
                  >
                    No users found.
                  </td>
                </tr>
              )}

            </tbody>

          </table>
        )}

      </div>
    </div>
  );
};

export default ManageUsers;