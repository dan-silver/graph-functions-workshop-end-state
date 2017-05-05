import { GraphClient } from '../authHelpers';
import { User } from '@microsoft/microsoft-graph-types' 

// Checked it worked in Graph explorer with https://graph.microsoft.com/beta/users?$select=displayName&$expand=extensions

export async function main (context, req) {
    if (context) context.log("Starting Azure function!");

    let users = await getUsers();

    for (let user of users) {

        // await removeExtensions(user);

        let numEvents = await queryNumberCalendarEvents(user);
        await saveUserExtension(user, Math.round(Math.random()*10));
    }

    let response = {
        status: 200,
        body: {
            // emails
        }
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

async function queryNumberCalendarEvents(user:User) {
    const client = await GraphClient();

    let today = new Date();
    let inOneMonth = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);

    return client
        .api(`/users/${user.mail}/calendarview/$count`)
        .query({
            startdatetime: today.toISOString(),
            enddatetime: inOneMonth.toISOString()
        })
        .get()
        .then((res) => {
            console.log(res)
            return res;
        })
        .catch((e) => {
            debugger;
        })
}


async function saveUserExtension(user:User, calendarEventsCount:number) {
    const client = await GraphClient();

    let extensionData = {
        extensionName: "numCalendarEvents",
        numEvents: calendarEventsCount
    };

    return client
        .api(`/users/${user.id}/extensions`)
        .version(`beta`)
        .post(extensionData)
        .catch((e) => {
            debugger;
        })
}

async function removeExtensions(user:User) {
    const client = await GraphClient();

    return client
        .api(`/users/${user.id}/extensions`)
        .version(`beta`)
        .get()
        .catch((e) => {
            debugger;
        }).then((res) => {
            let extensionIds = res['value'].map((extension) => extension.id);
            let extensionRemovals = [];
            for (let id of extensionIds) {
                extensionRemovals.push(
                        client
                        .api(`/users/${user.id}/extensions/${id}`)
                        .version(`beta`)
                        .delete()
                        .catch((e) => {
                            debugger;
                        }));
            }
            return Promise.all(extensionRemovals);
        })
} 

main(null, null);