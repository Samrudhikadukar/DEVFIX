import { SPFI, spfi, SPFx } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

import { WebPartContext } from "@microsoft/sp-webpart-base";

export class SharePointService {

  private sp: SPFI;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  // =========================================================
  // GET ALL SOLUTIONS
  // =========================================================

  public async getSolutions(): Promise<any[]> {

    const items = await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .select(
        "Id",
        "Title",
        "Technology",
        "Category",
        "Difficulty",
        "Problem",
        "ErrorMessage",
        "Cause",
        "Solution",
        "LessonLearned",
        "Tags",
        "HelpfulCount",
        "IsFavorite",
        "Status",
        "Created",
        "Author/Id",
        "Author/Title",
        "Author/EMail"
      )
      .expand("Author")();
 console.log("DEVFIX SOLUTIONS FROM SHAREPOINT:", items);
    return items;
  }

  // =========================================================
  // GET SINGLE SOLUTION
  // =========================================================

  public async getSolutionById(id: number): Promise<any> {

    const item = await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .getById(id)
      .select(
        "Id",
        "Title",
        "Technology",
        "Category",
        "Difficulty",
        "Problem",
        "ErrorMessage",
        "Cause",
        "Solution",
        "LessonLearned",
        "Tags",
        "HelpfulCount",
        "IsFavorite",
        "Status",
        "Created",
        "Author/Id",
        "Author/Title",
        "Author/EMail"
      )
      .expand("Author")();

    return item;
  }

  // =========================================================
  // GET MY CONTRIBUTIONS
  // =========================================================

  public async getMyContributions(
    userId: number
  ): Promise<any[]> {

    const items = await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .select(
        "Id",
        "Title",
        "Technology",
        "Category",
        "Difficulty",
        "Problem",
        "ErrorMessage",
        "Cause",
        "Solution",
        "LessonLearned",
        "Tags",
        "HelpfulCount",
        "IsFavorite",
        "Status",
        "Created",
        "Author/Id",
        "Author/Title",
        "Author/EMail"
      )
      .expand("Author")
      .filter(`Author/Id eq ${userId}`)
      .orderBy("Created", false)();

    return items;
  }

  // =========================================================
  // MARK SOLUTION AS HELPFUL
  // =========================================================

  public async markHelpful(id: number): Promise<void> {

    const item = await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .getById(id)
      .select("HelpfulCount")();

    const currentCount = item.HelpfulCount || 0;

    await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .getById(id)
      .update({
        HelpfulCount: currentCount + 1
      });
  }

  // =========================================================
  // CHECK FAVORITE
  // =========================================================

  public async isFavorite(
    solutionId: number,
    userId: number
  ): Promise<boolean> {

    const items = await this.sp.web.lists
      .getByTitle("DevFixFavorites")
      .items
      .filter(`SolutionId eq ${solutionId} and UserId eq ${userId}`)
      .top(1)();

    return items.length > 0;
  }

  // =========================================================
  // ADD FAVORITE
  // =========================================================

  public async addFavorite(
    solutionId: number,
    userId: number
  ): Promise<void> {

    await this.sp.web.lists
      .getByTitle("DevFixFavorites")
      .items
      .add({
        Title: `Solution ${solutionId}`,
        SolutionId: solutionId,
        UserId: userId
      });
  }

  // =========================================================
  // ADD SOLUTION
  // =========================================================

  public async addSolution(solution: {
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
    status: string;
  }): Promise<void> {

    await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .add({
        Title: solution.title,
        Technology: solution.technology,
        Category: solution.category,
        Difficulty: solution.difficulty,
        Problem: solution.problem,
        ErrorMessage: solution.errorMessage,
        Cause: solution.cause,
        Solution: solution.solution,
        LessonLearned: solution.lessonLearned,
        Tags: solution.tags,
        HelpfulCount: 0,
        IsFavorite: false,
        Status: solution.status
      });
  }

  // =========================================================
  // UPDATE SOLUTION
  // ADMIN / OWNER EDIT
  // =========================================================

  public async updateSolution(
    id: number,
    solution: {
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
  ): Promise<void> {

    await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .getById(id)
      .update({
        Title: solution.title,
        Technology: solution.technology,
        Category: solution.category,
        Difficulty: solution.difficulty,
        Problem: solution.problem,
        ErrorMessage: solution.errorMessage,
        Cause: solution.cause,
        Solution: solution.solution,
        LessonLearned: solution.lessonLearned,
        Tags: solution.tags
      });
  }



  public async getChoiceValues(
  columnName: string
): Promise<string[]> {

  const field = await this.sp.web.lists
    .getByTitle("DevFixSolutions")
    .fields
    .getByInternalNameOrTitle(columnName)();

  return field.Choices || [];
}
  // =========================================================
  // APPROVE SOLUTION
  // Pending Review -> Published
  // =========================================================

  public async approveSolution(id: number): Promise<void> {

    await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .getById(id)
      .update({
        Status: "Published"
      });
  }

  // =========================================================
  // REJECT SOLUTION
  // Pending Review -> Rejected
  // =========================================================

  public async rejectSolution(id: number): Promise<void> {

    await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .getById(id)
      .update({
        Status: "Rejected"
      });
  }

  // =========================================================
  // DELETE SOLUTION
  // =========================================================

  public async deleteSolution(id: number): Promise<void> {

    await this.sp.web.lists
      .getByTitle("DevFixSolutions")
      .items
      .getById(id)
      .delete();
  }

  // =========================================================
  // REMOVE FAVORITE
  // =========================================================

  public async removeFavorite(
    solutionId: number,
    userId: number
  ): Promise<void> {

    const items = await this.sp.web.lists
      .getByTitle("DevFixFavorites")
      .items
      .filter(`SolutionId eq ${solutionId} and UserId eq ${userId}`)
      .top(1)();

    if (items.length > 0) {

      await this.sp.web.lists
        .getByTitle("DevFixFavorites")
        .items
        .getById(items[0].Id)
        .delete();
    }
  }

  // =========================================================
// UPDATE SOLUTION STATUS
// =========================================================

public async updateSolutionStatus(
  id: number,
  status:
    | "Published"
    | "Pending Review"
    | "Draft"
    | "Rejected"
): Promise<void> {

  await this.sp.web.lists
    .getByTitle("DevFixSolutions")
    .items
    .getById(id)
    .update({
      Status: status
    });
}

public async getFavoriteSolutions(userId: number): Promise<any[]> {
  const favoriteItems = await this.sp.web.lists
    .getByTitle("DevFixFavorites")
    .items
    .select("Id", "SolutionId", "UserId")
    .filter(`UserId eq ${userId}`)
    .top(5000)();

  const favoriteIds = favoriteItems.map(
    (item: any) => Number(item.SolutionId)
  );

  if (favoriteIds.length === 0) {
    return [];
  }

  const allSolutions = await this.getSolutions();

  const favoriteSolutions = allSolutions.filter(
    (solution: any) => {
      const solutionId = Number(solution.Id);

      return (
        favoriteIds.indexOf(solutionId) !== -1 &&
        solution.Status === "Published"
      );
    }
  );

  return favoriteSolutions.map((solution: any) => ({
    ...solution,
    IsFavorite: true
  }));
}










// public async seedDevFixSolutions(): Promise<number> {
//   const list = this.sp.web.lists.getByTitle("DevFixSolutions");

//   const samples = [
//     ["React component not rendering", "React", "Debugging", "Low", "Component page par show nahi hota.", "Blank component", "Incorrect import ya rendering condition.", "Import/export names check karo aur JSX rendering condition verify karo.", "Component tree aur browser console check karna seekho.", "React, JSX"],
//     ["useEffect runs repeatedly", "React", "Performance", "High", "Effect repeated API calls karta hai.", "Repeated requests", "Incorrect dependency array.", "Dependencies review karo aur unnecessary state updates avoid karo.", "Effects ki dependencies carefully manage karo.", "React, Hooks"],
//     ["React state shows old value", "React", "Debugging", "Medium", "State setter ke baad old value dikhti hai.", "Old state in console", "State setter current render ka variable immediately change nahi karta.", "Previous state ke liye functional state update use karo.", "State updates scheduled hote hain.", "React, State"],
//     ["API data missing from table", "JavaScript", "API", "Medium", "API data milta hai par table empty hai.", "Empty table", "Response array ka path galat hai.", "Network tab mein response check karo aur correct array ko state mein set karo.", "API response structure verify karo.", "JavaScript, API"],
//     ["SPFx web part not appearing", "SPFx", "Deployment", "High", "Web part SharePoint page par nahi dikhta.", "Web part missing", "Package deployment ya manifest issue.", "App Catalog deployment, manifest aur browser console verify karo.", "Package aur page configuration check karo.", "SPFx, Deployment"],
//     ["SharePoint list returns 403", "SharePoint", "Permissions", "High", "App list data access nahi kar pata.", "403 Forbidden", "User ke paas list permission nahi hai.", "List permissions aur required API permissions verify karo.", "Site access aur list access alag ho sakte hain.", "SharePoint, Permissions"],
//     ["PnPjs items not loading", "PnPjs", "Debugging", "Medium", "PnPjs list items retrieve nahi karta.", "Request failure", "Incorrect SPFx context ya list title.", "PnPjs initialization, list title aur request URL verify karo.", "Context aur permissions check karna zaroori hai.", "PnPjs, SPFx"],
//     ["SharePoint internal field name mismatch", "SharePoint", "Configuration", "Medium", "API visible column name se field nahi dhoondh pati.", "Field not found", "Internal name display name se alag hai.", "Field ka InternalName check karke API query mein use karo.", "API mein correct internal field name use karo.", "SharePoint, Fields"],
//     ["API returns 401 Unauthorized", "JavaScript", "Authentication", "High", "API authentication fail hoti hai.", "401 Unauthorized", "Token missing ya expired hai.", "Valid access token aur Authorization header verify karo.", "Token target API ke liye valid hona chahiye.", "API, Authentication"],
//     ["API returns 404 Not Found", "JavaScript", "Debugging", "Low", "API endpoint nahi milta.", "404 Not Found", "URL ya resource ID galat hai.", "Endpoint URL, route, HTTP method aur resource verify karo.", "API route carefully check karo.", "API, Routing"],
//     ["CORS request blocked", "JavaScript", "Configuration", "High", "Browser cross-origin API request block karta hai.", "CORS policy error", "Server origin allow nahi karta.", "API server par required frontend origin aur headers allow karo.", "CORS server-side configuration se control hota hai.", "JavaScript, CORS"],
//     ["Access token expires", "JavaScript", "Authentication", "High", "API calls kuch time baad fail hoti hain.", "401 Unauthorized", "Access token expire ho gaya.", "Supported authentication library se fresh token acquire karo.", "Token expiry ko handle karna zaroori hai.", "API, Token"],
//     ["Azure Function returns 500", "Azure", "Debugging", "High", "Function request process nahi kar pati.", "500 Internal Server Error", "Unhandled exception ya missing configuration.", "Function logs check karo aur configuration validate karo.", "Logs root cause identify karne mein help karte hain.", "Azure, Functions"],
//     [".NET JSON deserialization failure", ".NET", "API", "Medium", "JSON request model mein convert nahi hoti.", "JsonException", "Payload aur model types match nahi karte.", "JSON properties, types aur Content-Type verify karo.", "Request data ko model ke saath match karo.", ".NET, JSON"],
//     ["HTTP trigger accepts invalid input", "Azure", "Configuration", "Medium", "Function invalid request process karti hai.", "Invalid data", "Input validation missing hai.", "Required fields validate karo aur invalid input par HTTP 400 return karo.", "Business logic se pehle validation karo.", "Azure, Validation"],
//     ["Git push rejected", "JavaScript", "Deployment", "Medium", "Local commits remote branch par push nahi hote.", "Push rejected", "Remote branch par new commits hain.", "Remote changes fetch karke team workflow ke according integrate karo.", "Push se pehle remote changes review karo.", "Git, GitHub"],
//     ["npm dependency installation fails", "TypeScript", "Configuration", "High", "npm dependencies install nahi hoti.", "ERESOLVE error", "Package versions incompatible hain.", "Peer dependencies aur Node.js version check karo; compatible versions use karo.", "Dependencies aur lockfile consistent rakho.", "npm, TypeScript"],
//     ["Mobile layout overflows", "React", "UI/UX", "Medium", "Mobile screen par UI cut ho jati hai.", "Horizontal scrolling", "Fixed widths ya missing responsive styles.", "Flexible widths, CSS Grid/Flexbox aur media queries use karo.", "Multiple screen sizes par UI test karo.", "React, Responsive"],
//     ["Power Automate approval gets stuck", "Power Automate", "Debugging", "High", "Approval flow complete nahi hota.", "Approval pending", "Approver ya flow condition issue.", "Run history, approver assignment aur connection status verify karo.", "Flow run history se troubleshooting shuru karo.", "Power Automate, Approval"],
//     ["Power Apps form does not save", "Power Apps", "Debugging", "Medium", "Form submit karne par SharePoint item save nahi hota.", "Submission failed", "Validation ya permission issue.", "Required fields, form mode, OnSuccess/OnFailure aur data source permissions check karo.", "Form validation aur save errors separately check karo.", "Power Apps, SharePoint"]
//   ];

//   const existingItems = await list.items
//     .select("Title")
//     .top(5000)();

//   const existingTitles = new Set(
//     existingItems.map((item: any) =>
//       String(item.Title).trim().toLowerCase()
//     )
//   );

//   let added = 0;

//   for (const s of samples) {
//     const title = s[0];

//     if (existingTitles.has(title.toLowerCase())) {
//       continue;
//     }

//     await list.items.add({
//       Title: s[0],
//       Technology: s[1],
//       Category: s[2],
//       Difficulty: s[3],
//       Problem: s[4],
//       ErrorMessage: s[5],
//       Cause: s[6],
//       Solution: s[7],
//       LessonLearned: s[8],
//       Tags: s[9],
//       HelpfulCount: 0,
//       IsFavorite: false,
//       Status: "Published"
//     });

//     existingTitles.add(title.toLowerCase());
//     added++;
//   }

//   return added;
// }
}