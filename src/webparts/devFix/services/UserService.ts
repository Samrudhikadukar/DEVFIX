import { SPFI, spfi, SPFx } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICurrentUser {
  id: number;
  name: string;
  email: string;
  initials: string;
  role: "USER" | "ADMIN";
}

export class UserService {
  private sp: SPFI;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  public async getCurrentUser(): Promise<ICurrentUser> {
    const user = await this.sp.web.currentUser();

    const adminGroups = await this.sp.web.currentUser.groups();

    const isAdmin = adminGroups.some(
      group => group.Title === "DevFix Admins"
    );

    const name = user.Title || "";
    const email = user.Email || "";

    const initials = name
      .split(" ")
      .filter(part => part.length > 0)
      .map(part => part.charAt(0))
      .join("")
      .substring(0, 2)
      .toUpperCase();

    return {
      id: user.Id,
      name,
      email,
      initials,
      role: isAdmin ? "ADMIN" : "USER"
    };
  }

  public async getUsers(): Promise<any[]> {
    const users = await this.sp.web.siteUsers
      .select(
        "Id",
        "Title",
        "Email",
        "LoginName",
        "PrincipalType"
      )();

    const adminGroups = await this.sp.web.siteGroups
      .getByName("DevFix Admins")
      .users
      .select("Id")();

    const adminIds: number[] = adminGroups.map(
      adminUser => adminUser.Id
    );

    const actualUsers = users.filter(
      user => user.PrincipalType === 1
    );

    return actualUsers.map(user => ({
      id: user.Id,
      name: user.Title || "",
      email: user.Email || "",
      loginName: user.LoginName || "",
      role: adminIds.indexOf(user.Id) !== -1
        ? "ADMIN"
        : "USER"
    }));
  }
}