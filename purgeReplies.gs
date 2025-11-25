/**
 * Primary function for the script. 
 * Deletes all replies (messages at index 1+) in a thread, keeping only the original (index 0).
 *
 * @author - dnagpal (Refactored for Pagination and Batching)
 */
function purgeTaggedReplies() {

  // 'newer_than:30d' effectively targets active threads. 
  // Adjust logic if you want to target old threads instead.
  const purgeQuery = { 'q': 'label:PurgeReplies newer_than:30d' };

  futurePurgeReplies('me', purgeQuery);
}


/**
 * Fetches threads matching the query, identifies replies (non-first messages), 
 * and batch deletes them. Handles Pagination.
 *
 * @param  {String} user  -  User's email address.
 * @param  {Object} query - String used to filter the Messages listed.
 */
function futurePurgeReplies(user, query) {
    let pageToken;
    let totalDeleted = 0;

    // PAGE LOOP: Ensure we process ALL threads, not just the first 100
    do {
        try {
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
                        
                        // Collect IDs of replies (skipping the first message)
                        let ids = getReplyIdsToDelete(messageList);
                        if (ids.length > 0) {
                            batchIdsToDelete = batchIdsToDelete.concat(ids);
                        }
                    } catch (e) {
                        Logger.log("ERROR processing Thread ID " + threadList.threads[i].id + ": " + e.message);
                    }
                }

                // BATCH DELETE
                if (batchIdsToDelete.length > 0) {
                    batchTrashMessages(user, batchIdsToDelete);
                    totalDeleted += batchIdsToDelete.length;
                }
            }

            pageToken = threadList.nextPageToken;

        } catch (e) {
            Logger.log("CRITICAL ERROR in Page Loop: " + e.message);
            break;
        }
        
    } while (pageToken);

    Logger.log(`Completed. Total replies moved to trash: ${totalDeleted}`);
}

/**
 * Fetches the list of messages matching a given thread ID.
 * Gmail API returns messages in chronological order by default.
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
 * Identifies all messages except the first one (Index 0).
 * Assumes Index 0 is the original email.
 * * @param {Object} messageList
 * @return {Array} ids - List of IDs to delete
 */
function getReplyIdsToDelete(messageList) {
    let idsToDelete = [];
    
    // Safety check: If there is 1 or 0 messages, there are no replies to delete.
    if (messageList.messages.length <= 1) {
        return idsToDelete;
    }

    // Start loop at 1 to SKIP the original message (Index 0)
    for (let i = 1; i < messageList.messages.length; i++) {
        idsToDelete.push(messageList.messages[i].id);
    }
    
    return idsToDelete;
}

/**
 * Trashes messages in batches of 1000 using batchModify.
 */
function batchTrashMessages(user, allIds) {
    const BATCH_SIZE = 1000;
    
    for (let i = 0; i < allIds.length; i += BATCH_SIZE) {
        let batch = allIds.slice(i, i + BATCH_SIZE);
        try {
            Logger.log(`Batch trashing ${batch.length} replies...`);
            Gmail.Users.Messages.batchModify({
                'ids': batch,
                'addLabelIds': ['TRASH'] 
            }, user);
        } catch (e) {
            Logger.log("ERROR in batch trash: " + e.message);
        }
    }
}