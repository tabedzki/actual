import { describe, expect, it, vi } from 'vitest';

import { resolveBuildRun } from './resolve.mjs';

const SHA = 'adf5d198282218f06ce0a6cbe847bee66f483332';
const OTHER_SHA = '4e5c0d5f87f1c4277b12a1811023a03db9ccb4b8';

function makeRun(id, overrides = {}) {
  return {
    id,
    head_sha: SHA,
    status: 'completed',
    conclusion: 'success',
    html_url: `https://github.com/actualbudget/actual/actions/runs/${id}`,
    ...overrides,
  };
}

function runsResponse(...runs) {
  return { data: { workflow_runs: runs } };
}

function artifactsResponse(...names) {
  return {
    data: { artifacts: names.map(name => ({ name, expired: false })) },
  };
}

function setup({
  listWorkflowRuns,
  listWorkflowRunArtifacts = vi.fn(),
  ...options
}) {
  const sleep = vi.fn().mockResolvedValue(undefined);
  const resolve = () =>
    resolveBuildRun({
      github: {
        rest: { actions: { listWorkflowRuns, listWorkflowRunArtifacts } },
      },
      owner: 'actualbudget',
      repo: 'actual',
      headSha: SHA,
      label: 'PR head',
      notFoundHint: 'build may still be running',
      timeoutMs: 1_000,
      sleep,
      ...options,
    });
  return { resolve, sleep, listWorkflowRuns, listWorkflowRunArtifacts };
}

describe('resolveBuildRun', () => {
  it('selects the exact commit and rejects a stale API result', async () => {
    const staleRun = makeRun(30768952310, { head_sha: OTHER_SHA });
    const expectedRun = makeRun(32157684269);
    const { resolve, sleep, listWorkflowRuns } = setup({
      listWorkflowRuns: vi
        .fn()
        .mockResolvedValue(runsResponse(staleRun, expectedRun)),
    });

    expect(await resolve()).toBe(expectedRun);
    expect(listWorkflowRuns).toHaveBeenCalledOnce();
    expect(sleep).not.toHaveBeenCalled();
  });

  it('does not filter the run list by status, event or branch', async () => {
    // GitHub's `status=success` filter has been seen omitting runs that
    // are completed/success (e.g. build run 42639183183), so the resolver
    // must list every run for the SHA and check the conclusion itself.
    const { resolve, listWorkflowRuns } = setup({
      listWorkflowRuns: vi.fn().mockResolvedValue(runsResponse(makeRun(1))),
    });

    await resolve();

    expect(listWorkflowRuns).toHaveBeenCalledWith({
      owner: 'actualbudget',
      repo: 'actual',
      workflow_id: 'build.yml',
      head_sha: SHA,
      per_page: 50,
    });
    const params = listWorkflowRuns.mock.calls[0][0];
    expect(params).not.toHaveProperty('status');
    expect(params).not.toHaveProperty('event');
    expect(params).not.toHaveProperty('branch');
  });

  it('prefers a successful run over a newer in-progress run', async () => {
    const inProgress = makeRun(2, { status: 'in_progress', conclusion: null });
    const succeeded = makeRun(1);
    const { resolve, sleep } = setup({
      listWorkflowRuns: vi
        .fn()
        .mockResolvedValue(runsResponse(inProgress, succeeded)),
    });

    expect(await resolve()).toBe(succeeded);
    expect(sleep).not.toHaveBeenCalled();
  });

  it('keeps polling while the build is still running', async () => {
    const running = makeRun(1, { status: 'in_progress', conclusion: null });
    const finished = makeRun(1);
    const { resolve, sleep, listWorkflowRuns } = setup({
      listWorkflowRuns: vi
        .fn()
        .mockResolvedValueOnce(runsResponse())
        .mockResolvedValueOnce(runsResponse(running))
        .mockResolvedValueOnce(runsResponse(finished)),
    });

    expect(await resolve()).toBe(finished);
    expect(listWorkflowRuns).toHaveBeenCalledTimes(3);
    expect(sleep).toHaveBeenCalledTimes(2);
  });

  it('fails immediately when the build finished without success', async () => {
    const failed = makeRun(42673059785, { conclusion: 'failure' });
    const { resolve, sleep, listWorkflowRuns } = setup({
      listWorkflowRuns: vi.fn().mockResolvedValue(runsResponse(failed)),
      timeoutMs: 60 * 60 * 1000,
    });

    await expect(resolve()).rejects.toThrow(
      /PR head build\.yml run 42673059785 finished with conclusion "failure".*actions\/runs\/42673059785/,
    );
    expect(listWorkflowRuns).toHaveBeenCalledOnce();
    expect(sleep).not.toHaveBeenCalled();
  });

  it('fails immediately for a cancelled build', async () => {
    const cancelled = makeRun(7, { conclusion: 'cancelled' });
    const { resolve, sleep } = setup({
      listWorkflowRuns: vi.fn().mockResolvedValue(runsResponse(cancelled)),
    });

    await expect(resolve()).rejects.toThrow(/conclusion "cancelled"/);
    expect(sleep).not.toHaveBeenCalled();
  });

  it('uses a failed build when every required artifact exists', async () => {
    // e.g. build run 42585252017 only failed in the mobile job, but all the
    // stats artifacts size-compare needs had been uploaded.
    const failed = makeRun(42585252017, { conclusion: 'failure' });
    const listWorkflowRunArtifacts = vi
      .fn()
      .mockResolvedValue(
        artifactsResponse('build-stats', 'api-build-stats', 'other'),
      );
    const log = vi.fn();
    const { resolve } = setup({
      listWorkflowRuns: vi.fn().mockResolvedValue(runsResponse(failed)),
      listWorkflowRunArtifacts,
      requiredArtifacts: ['build-stats', 'api-build-stats'],
      log,
    });

    expect(await resolve()).toBe(failed);
    expect(listWorkflowRunArtifacts).toHaveBeenCalledWith({
      owner: 'actualbudget',
      repo: 'actual',
      run_id: 42585252017,
      per_page: 100,
    });
    expect(log).toHaveBeenCalledWith(
      expect.stringContaining('has every required artifact'),
    );
  });

  it('fails when a failed build is missing a required artifact', async () => {
    const failed = makeRun(9, { conclusion: 'failure' });
    const { resolve } = setup({
      listWorkflowRuns: vi.fn().mockResolvedValue(runsResponse(failed)),
      listWorkflowRunArtifacts: vi
        .fn()
        .mockResolvedValue(artifactsResponse('build-stats')),
      requiredArtifacts: ['build-stats', 'cli-build-stats'],
    });

    await expect(resolve()).rejects.toThrow(/missing: cli-build-stats/);
  });

  it('treats expired artifacts as missing', async () => {
    const failed = makeRun(9, { conclusion: 'failure' });
    const { resolve } = setup({
      listWorkflowRuns: vi.fn().mockResolvedValue(runsResponse(failed)),
      listWorkflowRunArtifacts: vi.fn().mockResolvedValue({
        data: { artifacts: [{ name: 'build-stats', expired: true }] },
      }),
      requiredArtifacts: ['build-stats'],
    });

    await expect(resolve()).rejects.toThrow(/missing: build-stats/);
  });

  it('does not fall back to artifacts when none are required', async () => {
    const failed = makeRun(9, { conclusion: 'failure' });
    const { resolve, listWorkflowRunArtifacts } = setup({
      listWorkflowRuns: vi.fn().mockResolvedValue(runsResponse(failed)),
      requiredArtifacts: [],
    });

    await expect(resolve()).rejects.toThrow(/conclusion "failure"/);
    expect(listWorkflowRunArtifacts).not.toHaveBeenCalled();
  });

  it('waits for a newer in-progress rerun instead of failing on an older failure', async () => {
    const rerun = makeRun(2, { status: 'queued', conclusion: null });
    const failed = makeRun(1, { conclusion: 'failure' });
    const succeeded = makeRun(2);
    const { resolve, sleep } = setup({
      listWorkflowRuns: vi
        .fn()
        .mockResolvedValueOnce(runsResponse(rerun, failed))
        .mockResolvedValueOnce(runsResponse(succeeded, failed)),
    });

    expect(await resolve()).toBe(succeeded);
    expect(sleep).toHaveBeenCalledOnce();
  });

  it('times out when no run for the commit ever appears', async () => {
    const now = vi.fn().mockReturnValueOnce(0).mockReturnValue(11);
    const { resolve } = setup({
      listWorkflowRuns: vi.fn().mockResolvedValue(runsResponse()),
      timeoutMs: 10,
      now,
    });

    await expect(resolve()).rejects.toThrow(
      /No successful build\.yml run found for PR head within/,
    );
  });

  it('times out when only stale runs are returned', async () => {
    const now = vi.fn().mockReturnValueOnce(0).mockReturnValue(11);
    const { resolve } = setup({
      listWorkflowRuns: vi
        .fn()
        .mockResolvedValue(runsResponse(makeRun(1, { head_sha: OTHER_SHA }))),
      timeoutMs: 10,
      now,
    });

    await expect(resolve()).rejects.toThrow(/No successful build\.yml run/);
  });

  it('rejects a missing commit SHA', async () => {
    const { resolve, listWorkflowRuns } = setup({
      listWorkflowRuns: vi.fn(),
      headSha: '',
    });

    await expect(resolve()).rejects.toThrow(/without a commit SHA/);
    expect(listWorkflowRuns).not.toHaveBeenCalled();
  });
});
