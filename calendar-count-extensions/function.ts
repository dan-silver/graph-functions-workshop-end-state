import { GraphClient } from '../authHelpers';
import { User } from '@microsoft/microsoft-graph-types' 
import { queryNumberCalendarEvents, saveUserExtension, removeAllExtensionsOnUser } from "../graph-helpers";

// Checked it worked in Graph explorer with https://graph.microsoft.com/beta/users?$select=displayName&$expand=extensions

export async function main (context?, req?) {
    if (context) context.log("Starting Azure function!");

    // GET /users
    let users = await getUsers();

    for (let user of users) {

        // If you need to start over, you can clear extensions by uncommenting this (and commenting the saveUserExtension call!)
        // removeAllExtensionsOnUser(user);

        // how many events are on their calendar next week?
        let numEvents = await queryNumberCalendarEvents(user);

        // POST to the user with the num of calendar events as an extension
        await saveUserExtension(user, numEvents);
    }

    let response = {
        status: 200,
        body: "Saved extensions on users!"
    };
    return response;
};

async function getUsers() {
    const client = await GraphClient();

    return client
        .api("/users")
        .get()
        .then((res) => {
            return res.value as User[];
        });
}

main();