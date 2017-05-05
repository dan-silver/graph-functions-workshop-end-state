import { GraphClient } from '../authHelpers';
import { User, Message } from '@microsoft/microsoft-graph-types' 


export async function main (context, req) {
    if (context) context.log("Starting Azure function!");

    let users = await getUsersWithExtensions();
    users = users.sort(sortUsers).slice(0, 10);

    sendMailReport(users);
};

async function getUsersWithExtensions():Promise<User[]> {
    const client = await GraphClient();

    return client
        .api("/users")
        .version("beta")
        .expand("extensions")
        .get()
        .then((res) => {
            return res.value;
        })
}


async function sendMailReport(users) {
    const client = await GraphClient();

    let emailString = ''

    for (let user of users) {
        if (user.extensions)
            emailString += "<tr><td>" + user.displayName + "</td><td>" + user.extensions[0].numEvents + " events next week</td></tr>"
    }

    let message:Message = {
        subject: "Report on employee calendars",
        toRecipients: [{
            emailAddress: {
                address: "dansil@microsoft.com"
            }
        }],
        body: {
            content: `
                <table>

                ${emailString}
                </table>
            `,
            contentType: "html"
        }
    }
    return await client
        .api("/users/admin@MOD789932.onmicrosoft.com/sendMail")
        .post({message})
        .then((res) => {
            console.log("Mail sent!")
        }).catch((error) => {
            debugger;
        });
}



function sortUsers(a,b) {
  if (!a.extensions || !b.extensions) return 0;
  if (a.extensions[0].numEvents > b.extensions[0].numEvents)
    return -1;
  if (a.extensions[0].numEvents < b.extensions[0].numEvents)
    return 1;
  return 0;
}

main(null, null);