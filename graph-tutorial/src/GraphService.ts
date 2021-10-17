import { Client, GraphRequestOptions, PageCollection, PageIterator } from '@microsoft/microsoft-graph-client';
import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { endOfWeek, startOfWeek } from 'date-fns';
import { zonedTimeToUtc } from 'date-fns-tz';
import { User, Event } from 'microsoft-graph';
import { Task } from './Task';

let graphClient: Client | undefined = undefined;

function ensureClient(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
    if (!graphClient) {
        graphClient = Client.initWithMiddleware({
            authProvider: authProvider
        });
    }

    return graphClient;
}

export async function getUser(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<User> {
    ensureClient(authProvider);

    // Return the /me API endpoint result as a User object
    const user: User = await graphClient!.api('/me')
        // Only retrieve the specific fields needed
        .select('displayName,mail,mailboxSettings,userPrincipalName')
        .get();

    return user;
}

export async function getUserWeekCalendar(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    timeZone: string): Promise<Event[]> {
    ensureClient(authProvider);

    // Generate startDateTime and endDateTime query params
    // to display a 7-day window
    const now = new Date();
    const startDateTime = zonedTimeToUtc(startOfWeek(now), timeZone).toISOString();
    const endDateTime = zonedTimeToUtc(endOfWeek(now), timeZone).toISOString();

    // GET /me/calendarview?startDateTime=''&endDateTime=''
    // &$select=subject,organizer,start,end
    // &$orderby=start/dateTime
    // &$top=50
    var response: PageCollection = await graphClient!
        .api('/me/calendarview')
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({ startDateTime: startDateTime, endDateTime: endDateTime })
        .select('subject,organizer,start,end')
        .orderby('start/dateTime')
        .top(25)
        .get();

    if (response["@odata.nextLink"]) {
        // Presence of the nextLink property indicates more results are available
        // Use a page iterator to get all results
        var events: Event[] = [];

        // Must include the time zone header in page
        // requests too
        var options: GraphRequestOptions = {
            headers: { 'Prefer': `outlook.timezone="${timeZone}"` }
        };

        var pageIterator = new PageIterator(graphClient!, response, (event) => {
            events.push(event);
            return true;
        }, options);

        await pageIterator.iterate();

        return events;
    } else {

        return response.value;
    }
}

export async function createEvent(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    newEvent: Event): Promise<Event> {
    ensureClient(authProvider);

    // POST /me/events
    // JSON representation of the new event is sent in the
    // request body
    return await graphClient!
        .api('/me/events')
        .post(newEvent);
}

export async function getToDoLists(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
    ensureClient(authProvider);

    // Generate startDateTime and endDateTime query params
    // to display a 7-day window
    const now = new Date();


    let response = await graphClient!.api('/sites/woptopetest.sharepoint.com,83368a09-6832-4a7a-bbbe-a0706d14eb96,f5185c43-c2eb-4162-b08d-6ebc2154002a/lists/ToDoList/items?expand=fields(select=TaskName2,Description,StartDate,DueDate,Status,User,User_x003a_Email)').get();

    return response.value;
}

export async function createTask(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    newTask: Task) {
    ensureClient(authProvider);

    const listItem = {
        fields: {
            TaskName2: newTask.taskName,
            Description: newTask.description,
            StartDate: newTask.startDate,
            DueDate: newTask.dueDate,
            Status: newTask.status
        }
    };

    return await graphClient!
        .api('/sites/woptopetest.sharepoint.com,83368a09-6832-4a7a-bbbe-a0706d14eb96,f5185c43-c2eb-4162-b08d-6ebc2154002a/lists/ToDoList/items')
        .post(listItem);

}

export async function deleteTask(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    taskId: string) {
    ensureClient(authProvider);

    return await graphClient!
        .api('/sites/woptopetest.sharepoint.com,83368a09-6832-4a7a-bbbe-a0706d14eb96,f5185c43-c2eb-4162-b08d-6ebc2154002a/lists/ToDoList/items/' + taskId)
        .delete();

}

export async function getTask(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    id: string) {
    ensureClient(authProvider);


    const response = await graphClient!
        .api('/sites/woptopetest.sharepoint.com,83368a09-6832-4a7a-bbbe-a0706d14eb96,f5185c43-c2eb-4162-b08d-6ebc2154002a/lists/ToDoList/items/' + id + '?expand=fields(select=TaskName2,Description,StartDate,DueDate,Status,User,User_x003a_Email)')
        .get();


    const task: Task = {
        id: id,
        taskName: response.fields.TaskName2,
        description: response.fields.Description,
        startDate: response.fields.StartDate,
        dueDate: response.fields.DueDate,
        status: response.fields.Status
    }


    return task;
}

export async function updateTask(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    task: Task) {
    ensureClient(authProvider);

        const taskToUpdate = {
            TaskName2: task.taskName,
            Description: task.description,
            StartDate: task.startDate,
            DueDate: task.dueDate,
            Status: task.status
        }

        return await graphClient!
            .api('/sites/woptopetest.sharepoint.com,83368a09-6832-4a7a-bbbe-a0706d14eb96,f5185c43-c2eb-4162-b08d-6ebc2154002a/lists/ToDoList/items/' + task.id + '/fields')
            .update(taskToUpdate)

    }



