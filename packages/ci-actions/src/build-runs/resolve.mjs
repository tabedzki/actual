const DEFAULT_TIMEOUT_MS = 30 * 60 * 1000;
const DEFAULT_POLL_INTERVAL_MS = 15_000;

const defaultSleep = ms => new Promise(resolve => setTimeout(resolve, ms));

/**
 * Find a usable `workflowId` run for exactly `headSha`.
 *
 * Runs are listed without a `status` filter: GitHub's `status=success`
 * filter has been observed omitting runs that are completed/success, which
 * made this poll until its deadline. The conclusion is checked here instead,
 * which also lets a terminal failure fail fast rather than waiting out the
 * timeout. A failed run is still used when it uploaded every artifact in
 * `requiredArtifacts` (e.g. only an unrelated job failed).
 */
export async function resolveBuildRun({
  github,
  owner,
  repo,
  headSha,
  label,
  notFoundHint,
  workflowId = 'build.yml',
  requiredArtifacts = [],
  timeoutMs = DEFAULT_TIMEOUT_MS,
  sleepMs = DEFAULT_POLL_INTERVAL_MS,
  now = Date.now,
  sleep = defaultSleep,
  log,
}) {
  if (!headSha) {
    throw new Error(`Cannot resolve ${label} build without a commit SHA.`);
  }

  const deadline = now() + timeoutMs;

  while (true) {
    const { data } = await github.rest.actions.listWorkflowRuns({
      owner,
      repo,
      workflow_id: workflowId,
      head_sha: headSha,
      per_page: 50,
    });
    const runs = data.workflow_runs.filter(
      candidate => candidate.head_sha === headSha,
    );

    const unexpectedRun = data.workflow_runs.find(
      candidate => candidate.head_sha !== headSha,
    );
    if (unexpectedRun) {
      log?.(
        `Ignoring ${label} build run ${unexpectedRun.id}: expected ${headSha}, got ${unexpectedRun.head_sha}.`,
      );
    }

    const successfulRun = runs.find(
      run => run.status === 'completed' && run.conclusion === 'success',
    );
    if (successfulRun) {
      log?.(
        `Found ${label} build run ${successfulRun.id} (${successfulRun.html_url})`,
      );
      return successfulRun;
    }

    const pendingRun = runs.find(run => run.status !== 'completed');

    // Every run for this commit has finished and none succeeded; another
    // poll can't change that, so decide now instead of waiting out the
    // deadline. Runs are listed newest first.
    if (!pendingRun && runs.length > 0) {
      const [latestRun] = runs;
      let missingArtifacts = [];
      if (requiredArtifacts.length > 0) {
        missingArtifacts = await findMissingArtifacts({
          github,
          owner,
          repo,
          runId: latestRun.id,
          requiredArtifacts,
        });
        if (missingArtifacts.length === 0) {
          log?.(
            `Using ${label} build run ${latestRun.id} (${latestRun.html_url}): it finished with conclusion "${latestRun.conclusion}" but has every required artifact.`,
          );
          return latestRun;
        }
      }
      const missingNote =
        missingArtifacts.length > 0
          ? ` and is missing: ${missingArtifacts.join(', ')}`
          : '';
      throw new Error(
        `${label} ${workflowId} run ${latestRun.id} finished with conclusion "${latestRun.conclusion}"${missingNote} — see ${latestRun.html_url}. Re-run it, then re-run this workflow.`,
      );
    }

    if (now() >= deadline) {
      throw new Error(
        `No successful ${workflowId} run found for ${label} within ${timeoutMs / 60_000} min — ${notFoundHint}.`,
      );
    }

    const waitingFor = pendingRun
      ? `run ${pendingRun.id} is ${pendingRun.status}`
      : 'no run yet';
    log?.(
      `No successful ${label} build run yet (${waitingFor}) — sleeping ${sleepMs / 1000}s.`,
    );
    await sleep(sleepMs);
  }
}

async function findMissingArtifacts({
  github,
  owner,
  repo,
  runId,
  requiredArtifacts,
}) {
  const { data } = await github.rest.actions.listWorkflowRunArtifacts({
    owner,
    repo,
    run_id: runId,
    per_page: 100,
  });
  const available = new Set(
    data.artifacts
      .filter(artifact => !artifact.expired)
      .map(artifact => artifact.name),
  );
  return requiredArtifacts.filter(name => !available.has(name));
}
