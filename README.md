# Microsoft-Org-Chart-Builder

Uses the Microsoft Graph API to pull down an org chart from your desired top level position and formats it in a way that can be imported into Draw.io. In the future I may add a built-in interface for drawing the org chart, but for now Draw.io is easy, flexible and free, so might as well use it. At the moment this is a very basic tool developed for a specific purpose, but hopefully it can be helpful to others as well. 

You don't need to create an application in Azure to use this! 

Simply go to [https://developer.microsoft.com/en-us/graph/graph-explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and authenticate, make a basic request like `https://graph.microsoft.com/v1.0/me` then click the `Access token` tab and copy your token. 

You can make a copy of `default_key.yaml` named `user_key.yaml` and save your token here, or enter it into the prompt directly. 

You can also create a `user_cfg.yaml` and adjust the `default_cfg.yaml` settings as you like.

The resulting file will be placed in `/src` - probably not a great place, so that could change later. 

Instructions for how to import your file into Draw.io will be included at the top. You can adjust `src/drawio_csv_template.txt` to your liking as well if you want to preserve formatting options for future runs.