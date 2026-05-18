/**
 * Builds the Office.js Word code that applies proofreading changes as tracked edits.
 * Called by the server after the LLM returns proofreading_changes.
 * Track changes are enabled before edits so every replacement appears as a redline.
 *
 * @param {Array<{ find: string, replace: string, reason: string }>} changes
 */
export function buildProofreadCode(changes) {
  return `
(async () => {
  await Word.run(async (ctx) => {
    // Enable track changes so all edits appear as redlines
    ctx.document.changeTrackingMode = "trackAll";
    await ctx.sync();

    const changes = ${JSON.stringify(changes, null, 4)};

    let applied = 0, skipped = 0;

    for (const change of changes) {
      if (!change.find || !change.replace || change.find === change.replace) {
        skipped++;
        continue;
      }

      try {
        const results = ctx.document.body.search(change.find, {
          matchCase: false,
          matchWholeWord: false,
        });
        results.load("items");
        await ctx.sync();

        if (results.items.length === 0) {
          skipped++;
          continue;
        }

        const range = results.items[0];

        // Attach a comment with the reason (sidebar annotation)
        try {
          range.insertComment(change.reason);
          await ctx.sync();
        } catch (_) {
          // insertComment not available in this API version — continue without comment
        }

        // Replace text — recorded as a tracked deletion + insertion
        range.insertText(change.replace, "Replace");
        await ctx.sync();
        applied++;
      } catch (_) {
        skipped++;
      }
    }

    console.log("Proofreading: " + applied + " change(s) applied, " + skipped + " skipped.");
    // Leave track changes ON so the user can review, accept, or reject
  });
})()`.trim();
}
