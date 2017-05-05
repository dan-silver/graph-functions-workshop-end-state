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
// Checked it worked in Graph explorer with https://graph.microsoft.com/beta/users?$select=displayName&$expand=extensions
function main(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        if (context)
            context.log("Starting Azure function!");
        let users = yield getUsers();
        for (let user of users) {
            // await removeExtensions(user);
            let numEvents = yield queryNumberCalendarEvents(user);
            yield saveUserExtension(user, Math.round(Math.random() * 10));
        }
        let response = {
            status: 200,
            body: {}
        };
        return response;
    });
}
exports.main = main;
;
function getUsers() {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        return client
            .api("/users")
            .get()
            .then((res) => {
            return res.value;
        });
    });
}
function queryNumberCalendarEvents(user) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
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
            console.log(res);
            return res;
        })
            .catch((e) => {
            debugger;
        });
    });
}
function saveUserExtension(user, calendarEventsCount) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
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
        });
    });
}
function removeExtensions(user) {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
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
                extensionRemovals.push(client
                    .api(`/users/${user.id}/extensions/${id}`)
                    .version(`beta`)
                    .delete()
                    .catch((e) => {
                    debugger;
                }));
            }
            return Promise.all(extensionRemovals);
        });
    });
}
main(null, null);
//# sourceMappingURL=function.js.map