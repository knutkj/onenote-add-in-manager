# Retrieving Original PR Description (First Comment) on GitHub

When a pull request’s description (the top-most comment) has been edited, GitHub
does not directly show the original content. However, you can recover the
initial markdown using GitHub’s APIs. Below are two methods: using the GraphQL
API via GitHub CLI, and using the Issues Timeline REST API. Both methods will
fetch the original text that was generated by Copilot before any edits.

## Method 1: Using GitHub CLI with GraphQL API

GitHub’s GraphQL API provides a field `userContentEdits` for issues and pull
requests. This field contains a history of edits (including the diff of
changes). By querying this, you can retrieve details of all edits made to the PR
description. The **first edit entry** will contain the changes from the original
content to the first edit – allowing you to reconstruct the original.

**Steps:**

1. **Authenticate GitHub CLI:** Ensure you are logged in with the GitHub CLI and
   have access to the repository (e.g. run `gh auth login` if not already logged
   in).

2. **Prepare the GraphQL query:** You will query the repository’s pull request
   for its `userContentEdits`. For example, the query below fetches the PR’s
   current body and the edit history (including the diff and editor info for
   each edit):

   ```graphql
   query ($owner: String!, $repo: String!, $pr: Int!) {
     repository(owner: $owner, name: $repo) {
       pullRequest(number: $pr) {
         body
         userContentEdits(first: 10) {
           nodes {
             editedAt
             editor {
               login
             }
             diff
           }
         }
       }
     }
   }
   ```

   This will list up to 10 edits (adjust if needed). The `diff` field contains
   the markdown changes for each edit. The earliest entry in `userContentEdits`
   corresponds to the first edit made (i.e. changes from the original Copilot
   text to the first edited version).

3. **Run the GraphQL query via CLI:** Use `gh api` to execute the query. For
   example:

   ```bash
   gh api graphql -F owner='YOUR_OWNER' -F repo='YOUR_REPO' -F pr=PR_NUMBER -f query='
   query($owner: String!, $repo: String!, $pr: Int!) {
     repository(owner: $owner, name: $repo) {
       pullRequest(number: $pr) {
         body
         userContentEdits(first: 10) {
           nodes {
             editedAt
             editor { login }
             diff
           }
         }
       }
     }
   }'
   ```

   Replace `YOUR_OWNER/ YOUR_REPO` with the repository and `PR_NUMBER` with the
   pull request number.

4. **Inspect the output:** The output will be JSON containing the pull request’s
   current `body` and an array of `userContentEdits`. Each edit has a `diff`
   showing changes. For example, an edit entry might look like:

   ```json
   {
     "editedAt": "2025-06-01T...Z",
     "editor": { "login": "username" },
     "diff": "@@ -1,4 +1,4 @@\n-Original line of text\n+Edited line of text\n..."
   }
   ```

   In the `diff`, lines prefixed with `-` were in the original content and `+`
   lines are in the edited content. Using this, you can reconstruct the original
   **Copilot-generated text** by applying the diff in reverse (or simply reading
   the `-` lines as the content that was removed). If there were multiple edits,
   you may need to apply each diff stepwise or look at the earliest diff for the
   original text.

**Note:** The GraphQL approach gives you the changes but requires interpreting
the diff. It’s very powerful if you want a full edit history. The original
markdown is essentially the current text minus the diff changes from the first
edit.

## Method 2: Using the Issues Timeline REST API

GitHub’s Issues Timeline API can directly provide the previous body content for
edited events. When a pull request description is edited, the timeline will
include an **“edited” event**. This event contains a `changes` object with the
original text.

**Steps:**

1. **Ensure proper permissions:** If the repository is private, you need
   appropriate access and an authenticated GitHub CLI session or token. For
   public repos, no auth may be required, but it’s recommended to be
   authenticated.

2. **Fetch timeline events via CLI:** Use the `issues/timeline` endpoint for the
   pull request (pull requests are treated as issues in the API). For example:

   ```bash
   gh api -H "Accept: application/vnd.github+json" "/repos/YOUR_OWNER/YOUR_REPO/issues/PR_NUMBER/timeline"
   ```

   The `Accept` header ensures you get the timeline events in the JSON format.
   This will return an array of timeline events (comments, commits, edits,
   merges, etc.) for that issue/PR.

3. **Find the edited event:** Look through the JSON output for an event with
   `"event": "edited"`. You can use a tool like `jq` or grep to filter this. For
   example, piping to `jq`:

   ```bash
   gh api -H "Accept: application/vnd.github+json" "/repos/YOUR_OWNER/YOUR_REPO/issues/PR_NUMBER/timeline" | jq '.[] | select(.event=="edited")'
   ```

   This will show the **edit events**. Each edited event payload includes a
   `changes` field. Specifically, for pull request/issue edits, you should see
   something like:

   ```json
   {
     "event": "edited",
     "actor": { ... },
     "commit_id": null,
     "created_at": "2025-06-01T...Z",
     "changes": {
       "body": {
         "from": "*(original Copilot-generated markdown content here)*"
       }
     },
     "issue": { ... }
   }
   ```

   The `"body.from"` field contains the **previous content** before the edit. In
   the **first** edited event (chronologically), `"body.from"` will be the text
   that was initially created by Copilot.

4. **Retrieve the original content:** Copy the string under `changes.body.from`.
   This is the original markdown description as it was **prior to the first
   manual edit**. You can now use this content as needed (for example, restore
   it or compare with the current description).

**Tip:** If the pull request description was edited multiple times, there will
be multiple `"edited"` events. The earliest one contains the original text in
`from`. Later events would show intermediate versions. Ensure you identify the
first edit event to get the Copilot text.

## Additional Notes and Workarounds

- **GitHub Web UI:** If you prefer a visual route, on the GitHub website you can
  open the PR and click the “edited” label next to the description. GitHub will
  show a dropdown of revision history. There you can manually view or copy the
  original text. This is essentially the same data we retrieved via the API (it
  shows diffs or past versions).

- **Audit Logs (Enterprise only):** If you are on GitHub Enterprise and have
  access to audit logs, edits to issue descriptions might be logged. However,
  the methods above are easier and available on public GitHub.

- **Limitations:** The GitHub CLI’s standard commands (`gh pr view`) will only
  show the latest description, not the history. The API methods above are needed
  because GitHub does not directly expose prior versions in the CLI or REST v3
  without the timeline. The GraphQL `userContentEdits` and timeline events are
  the recommended ways to get historical content programmatically.

By following the steps in either method, you should be able to recover the
original markdown content that GitHub Copilot generated for the pull request
description before any edits were made. This allows you to review or restore the
initial PR comment as needed.

**Sources:**

- GitHub GraphQL API documentation and examples – showing how to query
  `userContentEdits` of a pull request.
- GitHub Issues timeline events API – the “edited” event includes a
  `changes.body.from` field containing the previous body text (original
  content).
