// H5-PowerBI-Connector

// This file contains your Data Connector logic
section H5Connector;

// TODO: Read c_id and c_sec from remote

// H5Connector OAuth2 values
client_id =  //client_ID from IONAPIFile; 
client_secret =  //client_secret from IONAPIFile; 
redirect_uri = "http://localhost:8080";
token_uri = //token_uri from IONAPIFile;
authorize_uri = //Oauth url from IONAPIFile;
logout_uri = "https://login.microsoftonline.com/logout.srf";

// Login modal window dimensions
windowWidth = 720;
windowHeight = 1024;

[DataSource.Kind="H5Connector", Publish="H5Connector.Publish"]
shared H5Connector.Contents = (H5_REST_Query_Url as text) =>
    let
        
        RestEndPoint = Json.Document(Web.Contents(H5_REST_Query_Url)),
        Results = RestEndPoint[results],
        NavResult = Results{0},
        Records = NavResult[records],
        TableStruct = Table.FromList(Records, Splitter.SplitByNothing(), null, null, ExtraValues.Error)
        //GetColumns = Table.ExpandRecordColumn(TableStruct, "Column1", {"DIVI","FACI","AGNB","PONR","IVAD","SAID","ITNO","BANO","LTYP","ORQT"},
        //{"Division","Facility","Agreement","Line","Invoice Address","Delivery Address","Item","Plant No","Line Type","Order Qty"}
        //)

       


    in
        TableStruct; 


// Data Source Kind description
H5Connector= [
  // TestConnection = () => { "H5Connector.Contents" },
    Authentication = [
        OAuth = [
            StartLogin=StartLogin,
            FinishLogin=FinishLogin,
            Refresh=Refresh,
            Logout=Logout
        ]
    ],
    Label = Extension.LoadString("DataSourceLabel")
];

// Data Source UI publishing description
H5Connector.Publish = [
    Beta = true,
    Category = "Other",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://powerbi.microsoft.com/",
    SourceImage = H5Connector.Icons,
    SourceTypeImage = H5Connector.Icons
];




// Helper functions for OAuth2: StartLogin, FinishLogin, Refresh, Logout
StartLogin = (resourceUrl, state, display) =>
    let
        authorizeUrl = authorize_uri & "?" & Uri.BuildQueryString([
            response_type = "code",
            client_id = client_id,  
            redirect_uri = redirect_uri
            //state = state,
            //scope = GetScopeString(scopes, scope_prefix)
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
        // parse the full callbackUri, and extract the Query string
        parts = Uri.Parts(callbackUri)[Query],
        // if the query string contains an "error" field, raise an error
        // otherwise call TokenMethod to exchange our code for an access_token
        result = if (Record.HasFields(parts, {"error", "error_description"})) then 
                    error Error.Record(parts[error], parts[error_description], parts)
                 else
                    TokenMethod("authorization_code", "code", parts[code])
    in
        result;

Refresh = (resourceUrl, refresh_token) => TokenMethod("refresh_token", "refresh_token", refresh_token);

Logout = (token) => logout_uri;


TokenMethod = (grantType, tokenField, code) =>
    let
        queryString = [
            grant_type = "authorization_code",
            redirect_uri = redirect_uri,
            client_id = client_id,
            client_secret = client_secret
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

        

H5Connector.Icons = [
    Icon16 = { Extension.Contents("H5Connector.png"), Extension.Contents("H5Connector.png"), Extension.Contents("H5Connector.png"), Extension.Contents("H5Connector.png") },
    Icon32 = { Extension.Contents("H5Connector.png"), Extension.Contents("H5Connector.png"), Extension.Contents("H5Connector.png"), Extension.Contents("H5Connector.png") }
];