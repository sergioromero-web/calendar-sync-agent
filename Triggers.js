function installTrigger() {
  uninstallPipelineTriggers();
  ScriptApp.newTrigger(TRIGGER_HANDLER)
    .timeBased()
    .everyHours(TRIGGER_EVERY_HOURS)
    .nearMinute(35)
    .create();
  Logger.log('Trigger installed for ' + TRIGGER_HANDLER + ' every ' + TRIGGER_EVERY_HOURS + ' hours at :35');
}

function uninstallPipelineTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === TRIGGER_HANDLER)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

function listPipelineTriggers() {
  return ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === TRIGGER_HANDLER)
    .map(trigger => ({
      handler: trigger.getHandlerFunction(),
      source: trigger.getTriggerSource(),
      eventType: trigger.getEventType()
    }));
}
