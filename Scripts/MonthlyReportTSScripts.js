/**
 * Gets statistics for all ad groups from a given period, sorted by label.
 * Statistics: ad group name, clicks, CPC, cost, label
 * @param period: string, period to fetch statistics from
 * @return: object of fetched statistics
 */
function getStatsFromAdGroups(period) {
  var stats = {};
  var noLabels = [];
  var adGroupIterator = AdWordsApp.adGroups().get();
  while (adGroupIterator.hasNext()) {
    var adGroup = adGroupIterator.next();
    var adGroupName = adGroup.getName();
    var campaignName = adGroup.getCampaign().getName();
    var adGroupStats = adGroup.getStatsFor(period);
    if (adGroupStats.getImpressions() > 0) {
      var labels = [];
      var labelsIterator = adGroup.labels().get();
      while (labelsIterator.hasNext()) {
        var label = labelsIterator.next().getName();
        if (label.substr(0,4).toLowerCase() !== "link") {labels.push(label);} // ignores LinkChecker labels
      }
      if (labels[0] === undefined && adGroupName !== "Android Games") {
        noLabels.push(adGroupName);
      }
      temp = {};
      temp[adGroupName + ";" + campaignName] = {"clicks": adGroupStats.getClicks(),
                                                "cpc": adGroupStats.getAverageCpc(),
                                                "cost": adGroupStats.getCost()};
      try {stats[labels[0]].push(temp);}
      catch (e) {
        stats[labels[0]] = [];
        stats[labels[0]].push(temp);
      }
    }
  }
  if (noLabels.length > 0) {
    var noLabelsString = "";
    for (var i = 0; i < noLabels.length; i++) {
      noLabelsString += noLabels[i] + "\n";
    }
    return noLabelsString;
  }
  return stats;
}

/**
 *
 */
function buildCSV(data) {
  var csv = "";
  for (var key1 in data) {
    for (var i = 0; i < data[key1].length; i++) {
      for (var key2 in data[key1][i]) {
        var split = key2.split(";");
        csv = csv + key1 + ";" + data[key1][i][key2]["clicks"] + ";" + data[key1][i][key2]["cpc"] + ";" + data[key1][i][key2]["cost"] + ";" + split[0] + ";" + split[1] + "\n";
      }

    }

  }
  return csv;
}

/*
 * @return: Date; date object for last day of previous month
 */
function getPrevMonth() {
  var previous = new Date();
  previous.setDate(1);
  previous.setHours(-1);
  return previous;
}

/**
 * Sends email to a specified recipient, with specified subject and message.
 * @param recipient: recipient of email
 * @param subject: subject of email
 * @param message: message of email
 */
function sendEmail(recipient, subject, body, message, attachment) {
  if (attachment) MailApp.sendEmail(recipient, subject, body, {attachments:[{fileName: attachment, mimeType: 'text/plain', content: message}]});
  else MailApp.sendEmail(recipient, subject, message);
}

function main() {
  var previous = getPrevMonth();
  var stats_data = getStatsFromAdGroups("LAST_MONTH");
  if (typeof stats_data === "string") {
    // sendEmail("damjan.mihelic@tsmedia.si", "GAdW Stats Report for Telekom, MISSING LABELS! ", "", "GAdW ad groups without labels:\n\n" + stats_data + "\nApply labels and rerun Monthly Report script.");
	  sendEmail("maja.cebulj@tsmedia.si", "GAdW Stats Report for Telekom, MISSING LABELS! ", "", "GAdW ad groups without labels:\n\n" + stats_data + "\nApply labels and rerun Monthly Report script.");
    sendEmail("alen.savic@tsmedia.si", "GAdW Stats Report for Telekom, MISSING LABELS! ", "", "GAdW ad groups without labels:\n\n" + stats_data + "\nApply labels and rerun Monthly Report script.");
    sendEmail("urska.grad@tsmedia.si", "GAdW Stats Report for Telekom, MISSING LABELS! ", "", "GAdW ad groups without labels:\n\n" + stats_data + "\nApply labels and rerun Monthly Report script.");
  } else {
    var csv_data = buildCSV(stats_data);
    // sendEmail("damjan.mihelic@tsmedia.si", "GAdW Stats Report for Telekom, Previous Month, CSV ", "", csv_data, "GAdW_Telekom_CSV_" + previous.getFullYear() + "-" + ((previous.getMonth() + 2 >= 10) ? previous.getMonth() + 2 : "0" + String(previous.getMonth() + 2)) + ".csv", 'text/csv');
    sendEmail("maja.cebulj@tsmedia.si", "GAdW Stats Report for Telekom, Previous Month, CSV ", "", csv_data, "GAdW_Telekom_CSV_" + previous.getFullYear() + "-" + ((previous.getMonth() + 2 >= 10) ? previous.getMonth() + 2 : "0" + String(previous.getMonth() + 2)) + ".csv", 'text/csv');
    sendEmail("alen.savic@tsmedia.si", "GAdW Stats Report for Telekom, Previous Month, CSV ", "", csv_data, "GAdW_Telekom_CSV_" + previous.getFullYear() + "-" + ((previous.getMonth() + 2 >= 10) ? previous.getMonth() + 2 : "0" + String(previous.getMonth() + 2)) + ".csv", 'text/csv');
    sendEmail("urska.grad@tsmedia.si", "GAdW Stats Report for Telekom, Previous Month, CSV ", "", csv_data, "GAdW_Telekom_CSV_" + previous.getFullYear() + "-" + ((previous.getMonth() + 2 >= 10) ? previous.getMonth() + 2 : "0" + String(previous.getMonth() + 2)) + ".csv", 'text/csv');
  }
}
