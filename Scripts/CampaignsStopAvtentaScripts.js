var monthlyBudget = 700;

/**
 * Pauses all campaings.
 * @param acc: account to pause campaigns within
 */
function pauseAllCampaigns() {
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    campaign = campaignIterator.next().pause();
  }
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
  var totalCost = 0;
  var adGroupIterator = AdWordsApp.adGroups().get();
  while (adGroupIterator.hasNext()) {
    var adGroup = adGroupIterator.next();
    var adGroupStats = adGroup.getStatsFor("THIS_MONTH");
    totalCost += adGroupStats.getCost();
  }
  if (totalCost >= monthlyBudget) {
    pauseAllCampaigns();
  	// sendEmail("damjan.mihelic@tsmedia.si", "Avtenta Monthly Budget Spent, Campaigns Stopped", "", "Avtenta Monthly Budget of €" + monthlyBudget + " Spent.\nTotal cost: €" + totalCost + "\nAll Campaigns Stopped.");
    sendEmail("maja.cebulj@tsmedia.si", "Avtenta Monthly Budget Spent, Campaigns Stopped", "", "Avtenta Monthly Budget of €" + monthlyBudget + " Spent.\nTotal cost: €" + totalCost + "\nAll Campaigns Stopped.");
    sendEmail("alen.savic@tsmedia.si", "Avtenta Monthly Budget Spent, Campaigns Stopped", "", "Avtenta Monthly Budget of €" + monthlyBudget + " Spent.\nTotal cost: €" + totalCost + "\nAll Campaigns Stopped.");
    sendEmail("urska.grad@tsmedia.si", "Avtenta Monthly Budget Spent, Campaigns Stopped", "", "Avtenta Monthly Budget of €" + monthlyBudget + " Spent.\nTotal cost: €" + totalCost + "\nAll Campaigns Stopped.");
  }
}
