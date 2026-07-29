function main() {
  const runId = Utilities.getUuid();
  const startedAt = new Date();
  const report = {
    runId: runId,
    startedAt: startedAt,
    finishedAt: null,
    sources: [],
    bulk: null,
    fatalError: null
  };

  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lockAcquired = lock.tryLock(LOCK_TIMEOUT_MS);
    if (!lockAcquired) {
      throw new Error('Another pipeline execution is still running');
    }

    Logger.log('Start pipeline runId=' + runId);

    const queueSheet = ensureExportQueue();
    reclaimStaleQueueRows(runId);
    reconcileSentQueueToState();
    pruneOldSentQueueRows(queueSheet);

    report.sources = runAll(runId);
    report.bulk = pushToBulkSafe(runId);

    // A second pass makes the state resilient if a previous run wrote to Bulk
    // but failed before updating the source state tabs.
    reconcileSentQueueToState();
    pruneOldSentQueueRows(queueSheet);

    report.finishedAt = new Date();
    notifyOps(report);
    Logger.log('End pipeline runId=' + runId);

  } catch (e) {
    report.finishedAt = new Date();
    report.fatalError = formatError(e);
    Logger.log('Pipeline ERROR runId=' + runId + ' | ' + report.fatalError);
    try {
      notifyOps(report);
    } catch (notifyError) {
      Logger.log('notifyOps ERROR: ' + formatError(notifyError));
    }
    throw e;

  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function pushToBulkSafe(runId) {
  const minutes = new Date().getMinutes();
  if (minutes >= 55 || minutes < 5) {
    Logger.log('pushToBulkSafe: unsafe window (' + minutes + 'min) — skipping');
    return { skipped: true, reason: 'unsafe window', minute: minutes };
  }
  return pushToBulk(runId);
}

function runAll(runId) {
  const summary = [];

  getActiveSources().forEach(cfg => {
    try {
      const result = syncSource(cfg, runId);
      summary.push(Object.assign({ key: cfg.key, success: true }, result));
    } catch (e) {
      const error = formatError(e);
      summary.push({ key: cfg.key, rows: 0, queued: 0, success: false, error: error });
      Logger.log('[' + cfg.key + '] ERROR: ' + error);
    }
  });

  return summary;
}

function getActiveSources() {
  return SOURCE.filter(cfg => cfg.active !== false);
}
