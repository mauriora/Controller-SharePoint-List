# Controller SharePoint list

Provides access to SharePoint lists.

- [Getting Started](#getting-started)
- [Examples](#examples)
  - [Announcements](#announcements)

## Getting Started

Not quite public yet, this is part of the [hybrid repro MVC SharePoint example implementation](https://github.com/mauriora/reusable-hybrid-repo-mvc-spfx-examples)

## Examples

### Announcements

This shows how get items from the Announcements List.

```typescript
    /** import the model */
    import { AnnouncementExtended } from '@mauriora/model-announcement-extended';

    /** import the controller factory */
    import { getCreateByIdOrTitle } from '@mauriora/controller-sharepoint-list';

    const newController = await getCreateByIdOrTitle(listName, siteUrl);

    const now: string = new Date().toISOString();
    /** get the SharePoint model */
    const newModel = await newController.addModel(
        AnnouncementExtended,
        `(StartDate le datetime'${now}' or StartDate eq null) and (Expires ge datetime'${now}' or Expires eq null)`
    );
    /** newModel.records is an Array of AnnouncementExtended */
    if(0 === newModel.records.length )
    {
        await newModel.loadAllRecords();
    }
```
