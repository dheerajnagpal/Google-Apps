/**
 * Primary function for the script. Change the label names to match your own setting 
 *
 * @author - dnagpal (Refactored for Pagination and Batching)
 * Note: Please use utility function at end to get label names.
 */
function deleteTaggedMessages() {

  const shortTermLabel = { 'q': 'label:7_Day_Delete older_than:7d' };
  const shortTermDays = 7;
  const mediumTermLabel = { 'q': 'label:30_Day_Delete older_than:30d' };
  const mediumTermDays = 30;

  // Add more rows if more label and day combinations are needed
  futureDeleteMessages('me', shortTermLabel, shortTermDays);
  futureDeleteMessages('me', mediumTermLabel, mediumTermDays);

}

/**
 * Fetches the threads tagged to a label, then delete messages that are older than the days parameter. 
 * Handles Pagination to ensure ALL threads are processed, not just the first 100.
 * Uses Batch operations to reduce API calls.
 *
 * @param  {String} user  -  User's email address. 
 * @param  {Object} query - String used to filter the Messages listed.
 * @param  {Number} numberOfDays - Number of days the message has to be older than today to be deleted.
 */
function futureDeleteMessages(user, query, numberOfDays) {
    let pageToken;
    let totalDeleted = 0;

    // PAGE LOOP: Keep fetching threads as long as there is a "next page"
    do {
        try {
            // Fetch threads with pageToken
            let threadList = Gmail.Users.Threads.list('me', {
                q: query.q,
                pageToken: pageToken
            });

            if (threadList.threads && threadList.threads.length > 0) {
                Logger.log(`Processing page with ${threadList.threads.length} threads...`);
                
                let batchIdsToDelete = [];

                // THREAD LOOP
                for (let i = 0; i < threadList.threads.length; i++) {
                    try {
                        let threadID = threadList.threads[i].id;
                        let messageList = getMessages(user, threadID);
                        
                        // Collect IDs instead of deleting immediately
                        let ids = getMessageIdsToDelete(messageList, numberOfDays);
                        if (ids.length > 0) {
                            batchIdsToDelete = batchIdsToDelete.concat(ids);
                        }
                    } catch (e) {
                        Logger.log("ERROR processing Thread ID " + threadList.threads[i].id + ": " + e.message);
                    }
                }

                // BATCH DELETE: Process all found IDs for this page
                if (batchIdsToDelete.length > 0) {
                    batchTrashMessages(user, batchIdsToDelete);
                    totalDeleted += batchIdsToDelete.length;
                }
            }

            // Get token for next page (undefined if no more pages)
            pageToken = threadList.nextPageToken;

        } catch (e) {
            Logger.log("CRITICAL ERROR in Page Loop: " + e.message);
            break; // Stop loop on critical API error
        }
        
    } while (pageToken);

    Logger.log(`Completed. Total messages moved to trash: ${totalDeleted}`);
}

/**
 * Fetches the list of messages matching a given thread ID.  
 */
function getMessages(user, threadID) {
    let thread = Gmail.Users.Threads.get(user, threadID);
    let messageList = { 'messages': [] };
    
    if (!thread.messages) return messageList;

    for (let j = 0; j < thread.messages.length; j++) {
        messageList.messages.push({
            'id': thread.messages[j].id,
            'messageDate': thread.messages[j].internalDate
        });
    }
    return messageList;
}

/**
 * Checks dates and returns a list of IDs that should be deleted.
 * * @return {Array} ids - Array of strings (message IDs)
 */
function getMessageIdsToDelete(messageList, numberOfDays) {
    let idsToDelete = [];
    let currentDate = new Date();
    let previousDate = new Date();
    previousDate.setDate(currentDate.getDate() - numberOfDays);
    let previousTime = previousDate.getTime();

    for (let i = 0; i < messageList.messages.length; i++) {
        if (messageList.messages[i].messageDate < previousTime) {
            idsToDelete.push(messageList.messages[i].id);
        }
    }
    return idsToDelete;
}

/**
 * Trashes messages in batches of 1000 (API Limit).
 * Note: batchModify with addLabelIds:['TRASH'] is the efficient way to "trash" items in bulk.
 */
function batchTrashMessages(user, allIds) {
    const BATCH_SIZE = 1000;
    
    for (let i = 0; i < allIds.length; i += BATCH_SIZE) {
        let batch = allIds.slice(i, i + BATCH_SIZE);
        try {
            Logger.log(`Batch trashing ${batch.length} messages...`);
            Gmail.Users.Messages.batchModify({
                'ids': batch,
                'addLabelIds': ['TRASH'] 
            }, user);
        } catch (e) {
            Logger.log("ERROR in batch trash: " + e.message);
        }
    }
}