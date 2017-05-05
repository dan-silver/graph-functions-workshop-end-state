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
const graph_helpers_1 = require("../graph-helpers");
// Checked it worked in Graph explorer with https://graph.microsoft.com/beta/users?$select=displayName&$expand=extensions
function main(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        if (context)
            context.log("Starting Azure function!");
        // GET /users
        let users = yield getUsers();
        for (let user of users) {
            // If you need to start over, you can clear extensions by uncommenting this (and commenting the saveUserExtension call!)
            // removeAllExtensionsOnUser(user);
            // how many events are on their calendar next week?
            let numEvents = yield graph_helpers_1.queryNumberCalendarEvents(user);
            // POST to the user with the num of calendar events as an extension
            yield graph_helpers_1.saveUserExtension(user, numEvents);
        }
        let response = {
            status: 200,
            body: "Saved extensions on users!"
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
main(null, null);
//# sourceMappingURL=function.js.map