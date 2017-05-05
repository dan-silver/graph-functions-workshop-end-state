"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const authHelpers_1 = require("../authHelpers");
function main(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        if (context)
            context.log("Starting Azure function!");
        let users = yield getUsersWithExtensions();
        users = users.sort(sortUsers).slice(0, 10);
        sendMailReport(users);
    });
}
exports.main = main;
;
function getUsersWithExtensions() {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        return client
            .api("/users")
            .version("beta")
            .expand("extensions")
            .get()
            .then((res) => {
            return res.value;
        });
    });
}
function sendMailReport(users) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        let emailString = '';
        for (let user of users) {
            if (user.extensions)
                emailString += "<tr><td>" + user.displayName + "</td><td>" + user.extensions[0].numEvents + " events next week</td></tr>";
        }
        let message = {
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
        };
        return yield client
            .api("/users/admin@MOD789932.onmicrosoft.com/sendMail")
            .post({ message })
            .then((res) => {
            console.log("Mail sent!");
        }).catch((error) => {
            debugger;
        });
    });
}
function sortUsers(a, b) {
    if (!a.extensions || !b.extensions)
        return 0;
    if (a.extensions[0].numEvents > b.extensions[0].numEvents)
        return -1;
    if (a.extensions[0].numEvents < b.extensions[0].numEvents)
        return 1;
    return 0;
}
main(null, null);
//# sourceMappingURL=function.js.map