/**
 * Primary function for the script. Delete all responses and just keep an original mail in the thread that matches a given label.
 * Change the label names to match your own setting 
 *
 * @author - dnagpal
 * TODO - next page token to be handled if the number of messages is too many. 
 * What would happen if a message thread ends up getting a next page token? Will it delete partial thread and rest will linger?
 * 
 * Note: Please use utility function getLables to get all label names in your account.
 */
function purgeTaggedReplies() {


  const purgeThirtyDate = { 'q': 'label:PurgeReplies newer_than:30d' };

  // enable the method for the duration for which to purge responses. 
    futurePurgeReplies('me', purgeThirtyDate);

}


/**
 * Fetches the threads tagged to a label, then delete messages that are older than the days parameter. 
 *
 * @param  {String} user  -  User's email address. The special value 'me' can be used to indicate the authenticated user.
 * @param  {Object} query - String used to filter the Messages listed.
 * @param  {Number} numberOfDays - Number of days the message has to be older than today to be deleted..
 */
function futurePurgeReplies(user, query) {

    // Get the threads matching the query.  
    let threads = getThreads(user, query);
    let threadID = null;
    let messageList = null;
    Logger.log("Total threads in search are: " + threads.resultSizeEstimate);
    for (let i = 0; i < threads.resultSizeEstimate; i++) {
        threadID = threads.threads[i].id;
        // For each thread, get the list of messages
        messageList = getMessages(user, threadID);
       Logger.log('Complete Message List is ' + JSON.stringify(messageList));
        // For the message list, delete the replies.
        deleteNonFirstMessages(user, messageList);

    }
}

/**
 * Fetches the threads matching a given query.  
 *
 * @param  {String} user  -  User's email address. The special value 'me' can be used to indicate the authenticated user.
 * @param  {Object} query - String used to filter the Messages listed.
 * @return {Object} threads - Object with list of threads. 
 */
function getThreads(user, query) {
    let threads = Gmail.Users.Threads.list('me', query);
    return threads;
}

/**
 * Fetches the list of messages matching a given thread ID.  
 *
 * @param  {String} user  -  User's email address. The special value 'me' can be used to indicate the authenticated user.
 * @param {String} threadID - Thread ID of the thread. 
 * @return {Object} messageList - Object with list of messages {ID and Message Date}
 */
function getMessages(user, threadID) {
    let thread = Gmail.Users.Threads.get(user, threadID);
    let messageList = {
        'messages': [
            { 'id': null, 'messageDate': null }
        ]
    };
    for (let j = 0; j < thread.messages.length; j++) {

        // Logger.log('Thread ID is ' + thread.messages[j].id);
        messageList.messages[j] = { 'id': null, 'messageDate': null };

        messageList.messages[j].id = thread.messages[j].id;
        messageList.messages[j].messageDate = thread.messages[j].internalDate;
        // Logger.log(thread.messages[j].internalDate);
    }

    return messageList;
}

/**
 * Deletes the list of messages in a list except the first one. 
 *
 * @param  {String} user  -  User's email address. The special value 'me' can be used to indicate the authenticated user.
 * @return {Object} messageList - Object with list of messages {ID and Message Date}
 */
function deleteNonFirstMessages(user, messageList) {

    Logger.log("Total number of messages being reviewed are: " + messageList.messages.length);
    //From second item onwards, delete the messages.
    for (let i = 1; i < messageList.messages.length; i++) {
        Logger.log(messageList.messages[i].id + ' will be deleted');
        Gmail.Users.Messages.trash(user,messageList.messages[i].id);
    }
}



