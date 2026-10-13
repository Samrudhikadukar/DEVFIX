
import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  DefaultButton
} from "@fluentui/react";
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
  const [sortField, setSortField] = useState<string>("name");
  const [isSortedDescending, setIsSortedDescending] =
    useState<boolean>(false);

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

  const sortedUsers = useMemo(() => {
    const result = [...users];

    result.sort((a, b) => {
      let comparison = 0;

      switch (sortField) {
        case "name":
          comparison = (a.name || "").localeCompare(
            b.name || "",
            undefined,
            { sensitivity: "base" }
          );
          break;

        case "email":
          comparison = (a.email || "").localeCompare(
            b.email || "",
            undefined,
            { sensitivity: "base" }
          );
          break;

        case "role":
          comparison = (a.role || "").localeCompare(b.role || "");
          break;

        case "solutions":
          comparison =
            getSolutionCount(a.id) - getSolutionCount(b.id);
          break;

        default:
          comparison = 0;
      }

      return isSortedDescending ? -comparison : comparison;
    });

    return result;
  }, [users, solutions, sortField, isSortedDescending]);

  const handleColumnClick = (
    _event?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ): void => {
    if (!column?.fieldName) {
      return;
    }

    if (sortField === column.fieldName) {
      setIsSortedDescending((previous) => !previous);
    } else {
      setSortField(column.fieldName);
      setIsSortedDescending(false);
    }
  };

  const columns: IColumn[] = [
    {
      key: "name",
      name: "User",
      fieldName: "name",
      minWidth: 150,
      maxWidth: 280,
      isResizable: true,
      isSorted: sortField === "name",
      isSortedDescending: sortField === "name" && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (user?: IUser) => (
        <span className={styles.userName}>
          {user?.name || "Unknown User"}
        </span>
      )
    },
    {
      key: "email",
      name: "Email",
      fieldName: "email",
      minWidth: 180,
      maxWidth: 340,
      isResizable: true,
      isSorted: sortField === "email",
      isSortedDescending: sortField === "email" && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (user?: IUser) => (
        <span className={styles.email}>
          {user?.email || "-"}
        </span>
      )
    },
    {
      key: "role",
      name: "Role",
      fieldName: "role",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      isSorted: sortField === "role",
      isSortedDescending: sortField === "role" && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (user?: IUser) => (
        <span
          className={`${styles.roleBadge} ${
            user?.role === "ADMIN"
              ? styles.adminBadge
              : styles.userBadge
          }`}
        >
          {user?.role === "ADMIN" ? "Admin" : "User"}
        </span>
      )
    },
    {
      key: "solutions",
      name: "Solutions",
      fieldName: "solutions",
      minWidth: 100,
      maxWidth: 160,
      isResizable: true,
      isSorted: sortField === "solutions",
      isSortedDescending:
        sortField === "solutions" && isSortedDescending,
      onColumnClick: handleColumnClick,
      onRender: (user?: IUser) => (
        <span className={styles.solutionCount}>
          {user ? getSolutionCount(user.id) : 0}
        </span>
      )
    }
  ];

  return (
    <div className={styles.manageUsers}>
      <div className={styles.headerSection}>
        <div className={styles.headerContent}>
          <h1 className={styles.pageTitle}>Manage Users</h1>
          <p className={styles.pageSubtitle}>
            Members of the DevFix knowledge base and their roles
          </p>
        </div>

        <DefaultButton
          text="Back to Admin Dashboard"
          iconProps={{ iconName: "Back" }}
          onClick={onBack}
          className={styles.backButton}
        />
      </div>

      <div className={styles.tableContainer}>
        {loading && (
          <div className={styles.message} role="status">
            Loading users...
          </div>
        )}

        {!loading && error && (
          <div className={styles.errorMessage} role="alert">
            {error}
          </div>
        )}

        {!loading && !error && (
          users.length > 0 ? (
            <div className={styles.detailsListWrapper}>
              <DetailsList
                items={sortedUsers}
                columns={columns}
                setKey="manageUsers"
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                isHeaderVisible={true}
                onShouldVirtualize={() => false}
                className={styles.usersDetailsList}
              />
            </div>
          ) : (
            <div className={styles.emptyMessage}>
              No users found.
            </div>
          )
        )}
      </div>
    </div>
  );
};

export default ManageUsers;

