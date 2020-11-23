/**
 * Primary function for the script. Change the label names to match your own setting 
 *
 * @author - dnagpal
 * TODO - next page token to be handled if the number of messages is too many. 
 * What would happen if a message thread ends up getting a next page token? Will it delete partial thread and rest will linger?
 * 
 * Note: Please use utility funcetion at end to get label names.
 */
function deleteTaggedMessages() {

  const shortTermLabel = { 'q': 'label:7DayDelete older_than:7d' };
  const shortTermDays = 7;
  const mediumTermLabel = { 'q': 'label:30DayDelete older_than:30d' };
  const mediumTermDays = 30;

  // Add more rows if more label and day combinations are needed
    futureDeleteMessages('me', shortTermLabel, shortTermDays);
    futureDeleteMessages('me', mediumTermLabel, mediumTermDays);

}


/**
 * Fetches the threads tagged to a label, then delete messages that are older than the days parameter. 
 *
 * @param  {String} user  -  User's email address. The special value 'me' can be used to indicate the authenticated user.
 * @param  {Object} query - String used to filter the Messages listed.
 * @param  {Number} numberOfDays - Number of days the message has to be older than today to be deleted..
 */
function futureDeleteMessages(user, query, numberOfDays) {

    // Get the threads matching the query.  
    let threads = getThreads(user, query);
    let threadID = null;
    let messageList = null;
    Logger.log("Total threads in search are: " + threads.resultSizeEstimate);
    for (let i = 0; i < threads.resultSizeEstimate; i++) {
        threadID = threads.threads[i].id;
        // For each thread, get the list of messages
        messageList = getMessages(user, threadID);
        // Logger.log('Complete Message List is ' + JSON.stringify(messageList));
        // For the message list, delete the messages that are older than current date - days.
        deleteMessages(user, messageList, numberOfDays);

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
 * Fetches the list of messages matching a given thread ID.  
 *
 * @param  {String} user  -  User's email address. The special value 'me' can be used to indicate the authenticated user.
 * @return {Object} messageList - Object with list of messages {ID and Message Date}
 * @param {Number} numberOfDays - Number of days older than which a message is to be deleted. 
 */
function deleteMessages(user, messageList, numberOfDays) {
    let currentDate = new Date();
    currentDate.setTime(Date.now());
    // Logger.log('current Date is ' + currentDate.toString());
    let previousDate = new Date();
    previousDate.setDate(currentDate.getDate() - numberOfDays);
    // Logger.log('Purge Date is ' + previousDate);
    // Logger.log('inside deleteMessage. messageList is ' + JSON.stringify(messageList));
    Logger.log("Total number of messages being reviewed are: " + messageList.messages.length);
    for (let i = 0; i < messageList.messages.length; i++) {
        let date = new Date();
        date.setTime(messageList.messages[i].messageDate);

        if (messageList.messages[i].messageDate < previousDate.getTime()) {
            Logger.log(messageList.messages[i].id + ' will be deleted');
            Gmail.Users.Messages.trash(user,messageList.messages[i].id);
        }
        else {
            Logger.log(messageList.messages[i].messageDate + ' will not be deleted');
        }
    }
}



