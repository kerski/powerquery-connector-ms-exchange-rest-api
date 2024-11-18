/*
    Author: John Kerski

    Description: This Power Query M code defines a custom connector for Microsoft Exchange Online using the Microsoft Graph API. 
    The connector retrieves various data from Exchange Online, such as user events and group calendar views, by making HTTP requests to the Microsoft Graph API endpoints. 
    It includes OAuth authentication to securely access the data, handling login, token refresh, and logout processes. 
    The connector also provides functions to list events and group calendar views, addressing the limitations of the native Exchange Online connector by offering detailed information on recurring events, occurrences, and modifications to occurrences. 
    Additionally, it defines helper functions for making API requests and processing JSON responses. 
    The connector is designed to be used within Power BI or other Power Query environments, providing a user-friendly interface and navigation table for accessing the data.
*/
[Version = "1.0.0"]
section MSGRAPHEXCHANGEONLINE;
// Pull Client ID from Source File
client_id = Text.FromBinary(Extension.Contents("client_id.user")); 
redirect_uri = "https://oauth.powerbi.com/views/oauthredirect.html";
token_uri = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token";
authorize_uri = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize";
logout_uri = "https://login.microsoftonline.com/logout.srf";

url= "https://graph.microsoft.com/beta/groups?$expand=members";

windowWidth = 720;
windowHeight = 1024;

scope_prefix = "https://graph.microsoft.com/";
scopes = {
    "User.Read",
    "Calendars.Read",
    "Calendars.Read.Shared",
    "Calendars.ReadBasic"
};

[DataSource.Kind="MSGRAPHEXCHANGEONLINE", Publish="MSGRAPHEXCHANGEONLINE.UI"]
shared MSGRAPHEXCHANGEONLINE.FunctionCatalog = () =>
let
    objects = #table(
        {"Name","Key","Data","ItemKind","ItemName","IsLeaf"},
        {
            /* Get Calendar Information */
            {"ListEvents","ListEvents",MSGRAPHEXCHANGEONLINE.ListEvents,"Function","ListEvents",true},
            {"ListGroupCalendarView","ListGroupCalendarView",MSGRAPHEXCHANGEONLINE.ListGroupCalendarView,"Function","ListGroupCalendarView",true},
            /* Groups */
            {"ListUsersDirectMemberships","ListUsersDirectMemberships",MSGRAPHEXCHANGEONLINE.ListUsersDirectMemberships,"Function","ListUsersDirectMemberships",true}
        }),
    navTable = Table.ForceToNavigationTable(objects,{"Key"},"Name","Data","ItemKind","ItemName","IsLeaf")
in
    navTable;

/*****************
    Calendar
*****************/

/*** ListEvents ***/
shared MSGRAPHEXCHANGEONLINE.ListEvents = Value.ReplaceType(
    ListEvents,
    ListEventsType
);

ListEventsType = type function ( 
    optional timezone as ( 
        type text meta [
            Documentation.FieldCaption = "timezone",
            Documentation.FieldDescription = "Use this to specify the time zone for start and end times in the response. If not specified, those time values are returned in UTC. Optional.",
            Documentation.SampleValues = {
                "Eastern Standard Time"
            }
        ]
    )    
) as table meta [
    Documentation.Name = "MSGRAPHEXCHANGEONLINE.ListEvents",
    Documentation.LongDescription = "Retrieve a list of events in a calendar. The calendar can be one for a user.", 
    Documentation.Examples = {
        [
            Code = "=MSGRAPHEXCHANGEONLINE.ListEvents(""288a6513-8516-4dc9-a790-3706510exxxx"")",
            Result = "If successful, this method returns a 200 OK response code and a collection of Event objects in the response body. https://learn.microsoft.com/en-us/graph/api/calendar-list-events?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

ListEvents = (optional timezone as text) =>
let
    // Determine the timezone parameter, defaulting to "UTC" if not provided
    timezoneParam = if timezone <> null then timezone else "UTC",
    // Set the request headers with the specified timezone
    params = [#"Prefer"= "outlook.timezone=""" & timezoneParam & """"],
    eventCount = try Number.From(Json.Document(GetJSON("/me/calendar/events?$count=true"))[#"@odata.count"]) otherwise 0,
    
    // Calculate the number of iterations needed (ceiling of eventCount / 1000)
    iterations = Number.RoundUp(eventCount / 1000),

    // Initialize an empty list to store results
    allResults = List.Generate(
        () => [i = 0, result = GetJSON("/me/calendar/events?$top=1000&$skip=0",params)],
        each [i] < iterations,
        each [i = [i] + 1],
        each [result = GetJSON("/me/calendar/events?$top=1000&$skip=" & Text.From([i] * 1000),params)]
    ),

    // Combine all results into a single table
    result = 
    if eventCount = 0 then #table({},{}) 
    else    
        let
            #"Converted to Table" = Table.FromList(allResults, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"result"}, {"result"}),
            #"Parsed JSON" = Table.TransformColumns(#"Expanded Column1",{},Json.Document),
            #"Expanded result" = Table.ExpandRecordColumn(#"Parsed JSON", "result", {"value"}, {"value"}),
            #"Expanded value" = Table.ExpandListColumn(#"Expanded result", "value")
        in
        #"Expanded value"
in
    result;      
/*** End ListEvents ***/

/*** List Group Calendar View ***/
shared MSGRAPHEXCHANGEONLINE.ListGroupCalendarView = Value.ReplaceType(
    ListGroupCalendarView,
    ListGroupCalendarViewType
);

ListGroupCalendarViewType = type function ( 
    optional groupId as ( 
        type text meta [
            Documentation.FieldCaption = "Group ID",
            Documentation.FieldDescription = "The object ID for the group",
            Documentation.SampleValues = {
                "288a6513-8516-4dc9-a790-3706510exxxx"
            }
        ]
    ),
    optional startDateTime as ( 
        type text meta [
            Documentation.FieldCaption = "startDateTime",
            Documentation.FieldDescription = "The start date and time of the time range, represented in ISO 8601 format.",
            Documentation.SampleValues = {
                "2019-11-08T19:00:00-08:00"
            }
        ]
    ),
    optional endDateTime as ( 
        type text meta [
            Documentation.FieldCaption = "endDateTime",
            Documentation.FieldDescription = "The end date and time of the time range, represented in ISO 8601 format.",
            Documentation.SampleValues = {
                "2019-11-08T20:00:00-08:00"
            }
        ]
    ),
    optional timezone as ( 
        type text meta [
            Documentation.FieldCaption = "timezone",
            Documentation.FieldDescription = "Use this to specify the time zone for start and end times in the response. If not specified, those time values are returned in UTC. Optional.",
            Documentation.SampleValues = {
                "Eastern Standard Time"
            }
        ]
    )            
) as table meta [
    Documentation.Name = "MSGRAPHEXCHANGEONLINE.ListGroupCalendarView",
    Documentation.LongDescription = "Get the occurrences, exceptions, and single instances of events in a calendar view defined by a time range, from the default calendar of a group.", 
    Documentation.Examples = {
        [
            Code = "=MSGRAPHEXCHANGEONLINE.ListGroupCalendarView(""288a6513-8516-4dc9-a790-3706510exxxx"",""2019-11-08T19:00:00-08:00"",""2019-11-08T20:00:00-08:00"",""Eastern Standard Time"")",
            Result = "If successful, this method returns a 200 OK response code and collection of event objects in the response body. https://learn.microsoft.com/en-us/graph/api/group-list-calendarview?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

ListGroupCalendarView = (optional groupId as text, optional startDateTime as text, optional endDateTime as text, optional timezone as text) =>
let
    // Determine the timezone parameter, defaulting to "UTC" if not provided
    timezoneParam = if timezone <> null then timezone else "UTC",
    // Set the request headers with the specified timezone
    params = [#"Prefer"= "outlook.timezone=""" & timezoneParam & """"],
    eventCount = try Number.From(Json.Document(GetJSON("/groups/" & groupId & "/calendarView?startDateTime=" & startDateTime & "&endDateTime=" & endDateTime & "&$count=true"))[#"@odata.count"]) otherwise 0,
    
    // Calculate the number of iterations needed (ceiling of eventCount / 1000)
    iterations = Number.RoundUp(eventCount / 1000),

    // Initialize an empty list to store results
    allResults = List.Generate(
        () => [i = 0, result = GetJSON("/groups/" & groupId & "/calendarView?&startDateTime=" & startDateTime & "&endDateTime=" & endDateTime & "&$top=1000&$skip=0",params)],
        each [i] < iterations,
        each [i = [i] + 1],
        each [result = GetJSON("/groups/" & groupId & "/calendarView?&startDateTime=" & startDateTime & "&endDateTime=" & endDateTime & "&$top=1000&$skip=" & Text.From([i] * 1000),params)]
    ),

    // Combine all results into a single table
    result = 
    if eventCount = 0 then #table({},{}) 
    else
        let 
        #"Converted to Table" = Table.FromList(allResults, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"result"}, {"result"}),
        #"Parsed JSON" = Table.TransformColumns(#"Expanded Column1",{},Json.Document),
        #"Expanded result" = Table.ExpandRecordColumn(#"Parsed JSON", "result", {"value"}, {"value"}),
        #"Expanded value" = Table.ExpandListColumn(#"Expanded result", "value")
        in
        #"Expanded value"
in
    result;    

/*** End List Group Calendar View ***/

/*****************
    Groups
*****************/

/*** ListUsersDirectMemberships ***/
shared MSGRAPHEXCHANGEONLINE.ListUsersDirectMemberships = Value.ReplaceType(
    ListUsersDirectMemberships,
    ListUsersDirectMembershipsType
);

ListUsersDirectMembershipsType = type function ( 
) as table meta [
    Documentation.Name = "MSGRAPHEXCHANGEONLINE.ListUsersDirectMemberships",
    Documentation.LongDescription = "Get groups, directory roles, and administrative units that the user is a direct member of. This operation is not transitive.", 
    Documentation.Examples = {
        [
            Code = "=MSGRAPHEXCHANGEONLINE.ListUsersDirectMemberships()",
            Result = "If successful, this method returns a collection of directoryObject objects in the response body. https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

ListUsersDirectMemberships = () =>
    let 
        result = GetJSON("/me/memberOf")             
    in
        result;


/**** Help Functions ****/
GetJSON = (optional relPath as text, optional additionalHeaders as record) =>
    let
        headers = if additionalHeaders <> null then Record.Combine({[#"Content-Type" = "application/json"],additionalHeaders}) else [#"Content-Type" = "application/json"],
        params = [Headers = headers, RelativePath = relPath],
        //Get Pages
        //Hard-coded Base URL because Power BI Service won't recognize parameter despite numerous attempts
        webSource = Web.Contents(
            "https://graph.microsoft.com/v1.0",
            params
        ),
        //Convert to Json
        result = Text.FromBinary(webSource)
    in
        result;
/**** End Helper Functions ****/


MSGRAPHEXCHANGEONLINE = [
    TestConnection = (dataSourcePath) => { "MSGRAPHEXCHANGEONLINE.FunctionCatalog" },
    Authentication = [
        OAuth = [
            StartLogin=StartLogin,
            FinishLogin=FinishLogin,
            Refresh=Refresh,
            Logout=Logout
        ]
    ],
    Label = "Microsoft Exchange REST API Custom Connector"
];

// Setup Microsoft Exchange Connector
MSGRAPHEXCHANGEONLINE.UI = [
    Beta = true,
    Category = "Other",
    ButtonText = {
        Extension.LoadString(
            "ButtonTitle"
        ),
        Extension.LoadString(
            "ButtonHelp"
        )
    },    
    LearnMoreUrl = "https://github.com/kerski/powerquery-connector-ms-exchange-rest-api/blob/main/README.md#functions-implemented",
    SourceImage = MSGRAPHEXCHANGEONLINE.Icons,
    SourceTypeImage = MSGRAPHEXCHANGEONLINE.Icons
];

MSGRAPHEXCHANGEONLINE.Icons = [
    Icon16 = { Extension.Contents("powerquery-connector-ms-exchange-rest-api16.png"), Extension.Contents("powerquery-connector-ms-exchange-rest-api20.png"), Extension.Contents("powerquery-connector-ms-exchange-rest-api24.png"), Extension.Contents("powerquery-connector-ms-exchange-rest-api32.png") },
    Icon32 = { Extension.Contents("powerquery-connector-ms-exchange-rest-api32.png"), Extension.Contents("powerquery-connector-ms-exchange-rest-api40.png"), Extension.Contents("powerquery-connector-ms-exchange-rest-api48.png"), Extension.Contents("powerquery-connector-ms-exchange-rest-api64.png") }
];


// Thanks to https://aycabas.com/2020/05/11/graph-api-custom-connector/ for the artcle I found to do
// OAuth with Graph API.
StartLogin = (resourceUrl, state, display) =>
    let
        authorizeUrl = authorize_uri & "?" & Uri.BuildQueryString([
            client_id = client_id,  
            redirect_uri = redirect_uri,
            state = state,
            scope = "offline_access " & GetScopeString(scopes, scope_prefix),
            response_type = "code",
            response_mode = "query",
            login = "login"
        ])
    in
        [
            LoginUri = authorizeUrl,
            CallbackUri = redirect_uri,
            WindowHeight = 720,
            WindowWidth = 1024,
            Context = null
        ];


FinishLogin = (context, callbackUri, state) =>
    let
        parts = Uri.Parts(callbackUri)[Query],
 
        result = if (Record.HasFields(parts, {"error", "error_description"})) then 
                    error Error.Record(parts[error], parts[error_description], parts)
                 else
                    TokenMethod("authorization_code", "code", parts[code])
    in
        result;

// Refresh
// Original Signature: https://learn.microsoft.com/en-us/power-query/handling-authentication#original-oauth-signatures
Refresh = (resourceUrl, refresh_token) => TokenMethod("refresh_token", "refresh_token", refresh_token);

// Logout
// Original Signature: https://learn.microsoft.com/en-us/power-query/handling-authentication#original-oauth-signatures
Logout = (token) => logout_uri;

// Token Method
// Refresh Token
TokenMethod = (grantType, tokenField, code) =>
    let
        queryString = [
            client_id = client_id,
            scope = "offline_access " & GetScopeString(scopes, scope_prefix),
            grant_type = grantType,
            redirect_uri = redirect_uri
        ],
        queryWithCode = Record.AddField(queryString, tokenField, code),

        tokenResponse = Web.Contents(token_uri, [
            Content = Text.ToBinary(Uri.BuildQueryString(queryWithCode)),
            Headers = [
                #"Content-type" = "application/x-www-form-urlencoded",
                #"Accept" = "application/json"
            ],
            ManualStatusHandling = {400} 
        ]),
        body = Json.Document(tokenResponse),
        result = if (Record.HasFields(body, {"error", "error_description"})) then 
                    error Error.Record(body[error], body[error_description], body)
                 else
                    body
    in
        result;

// Handle scope for OAuth
Value.IfNull = (a, b) => if a <> null then a else b;

GetScopeString = (scopes as list, optional scopePrefix as text) as text =>
    let
        prefix = Value.IfNull(scopePrefix, ""),
        addPrefix = List.Transform(scopes, each prefix & _),
        asText = Text.Combine(addPrefix, " ")
    in
        asText;

//To Navigation Table
Table.ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(
            table
        ),
        newTableType = Type.AddTableKey(
            tableType,
            keyColumns,
            true
        ) meta [
            NavigationTable.NameColumn = nameColumn,
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn,
            Preview.DelayColumn = itemNameColumn,
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(
            table,
            newTableType
        )
    in
        navigationTable;

// Force Navigation Table
Table.ForceToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(
            table
        ),
        newTableType = Type.AddTableKey(
            tableType,
            keyColumns,
            true
        ) meta [
            NavigationTable.NameColumn = nameColumn,
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn,
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(
            table,
            newTableType
        )
    in
        navigationTable;