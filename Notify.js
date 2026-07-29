function notifyOps(report) {
  const failedSources = (report.sources || []).filter(result => !result.success);
  const bulkFailed = report.bulk && (report.bulk.error || report.bulk.failed);
  const fatalError = report.fatalError;
  const hasProblems = Boolean(fatalError || bulkFailed || failedSources.length);

  if (!hasProblems && !NOTIFY_ON_SUCCESS) {
    Logger.log('notifyOps: pipeline completed without alertable errors');
    return;
  }

  const subject = hasProblems
    ? '[Bulk availability] Pipeline error runId=' + report.runId
    : '[Bulk availability] Pipeline OK runId=' + report.runId;
  const body = buildOpsMessage(report, failedSources, hasProblems);

  Logger.log('notifyOps: ' + subject + '\n' + body);

  if (!SEND_OPS_EMAIL || !OPS_EMAIL) return;

  try {
    MailApp.sendEmail({
      to: OPS_EMAIL,
      subject: subject,
      body: body
    });
  } catch (e) {
    Logger.log('MailApp.sendEmail ERROR: ' + formatError(e));
  }
}

function buildOpsMessage(report, failedSources, hasProblems) {
  const lines = [];

  lines.push('Run Id: ' + report.runId);
  lines.push('Status: ' + (hasProblems ? 'ERROR' : 'OK'));
  lines.push('Started: ' + report.startedAt);
  lines.push('Finished: ' + report.finishedAt);

  if (report.fatalError) {
    lines.push('');
    lines.push('Fatal error:');
    lines.push(truncateText(report.fatalError, 4000));
  }

  if (report.bulk) {
    lines.push('');
    lines.push('Bulk:');
    lines.push(JSON.stringify(report.bulk, null, 2));
  }

  if ((report.sources || []).length) {
    lines.push('');
    lines.push('Sources:');
    (report.sources || []).forEach(result => {
      lines.push(JSON.stringify(result, null, 2));
    });
  }

  if (failedSources.length) {
    lines.push('');
    lines.push('Failed sources: ' + failedSources.map(result => result.key).join(', '));
  }

  return lines.join('\n');
}
