## SPFx Web Part. Getting presence information of users from Microsoft Teams

Microsoft Teams was released in 2017, and it took about three years to get an API that makes it possible to get user online presence status. We had to use a workaround. But as for today, we have the [REST API method to get user presence information](https://docs.microsoft.com/en-us/graph/api/presence-get?view=graph-rest-beta&tabs=http){target=_blank}.

### Building the code

```bash
git clone https://github.com/vzhukov/spfx-teams-presence.git
npm install
gulp serve --nobrowser
```

### Preview the web part

Open your SharePoint Online workbench at:

https://[tenant].sharepoint.com/_layouts/15/workbench.aspx

and add the web part to the page.

### Post

Original blog post: https://blog.vitalyzhukov.ru/en/spfx-teams-presence-status-microsoft-graph
