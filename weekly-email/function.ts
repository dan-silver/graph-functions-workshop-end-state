import { GraphClient } from '../authHelpers';
import { User, Message } from '@microsoft/microsoft-graph-types'


import { getUsersWithExtensions, sortUsersOnNumCalEvents, sendMailReport } from "../graph-helpers";


export async function main (context, req) {
    if (context) context.log("Starting Azure function!");

    // GET /beta/users&$expand=extensions
    let users = await getUsersWithExtensions();

    // sort descending order of busiest calendars
    users.sort(sortUsersOnNumCalEvents)
    
    // get top 10 users with busiest calendars
    users = users.slice(0, 10);

    sendMailReport(users);
};

main(null, null);