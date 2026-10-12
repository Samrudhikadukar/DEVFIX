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
}