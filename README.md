# Controller SharePoint list

Provides access to SharePoint lists.

## Getting Started

Not quite public yet, this is part of the [hybrid repro MVC SharePoint example implementation](https://github.com/mauriora/reusable-hybrid-repo-mvc-spfx-examples)

## Examples

### Announcements

This shows how get items from the Announcements List.

1. create a SharePoint List controller using `getCreateByIdOrTitle`
2. get the SharePoint model using `newController.addModel( ModelClass, filterQuery );`
3. if no records have been loaded, call `newModel.loadAllRecords()`

```typescript
    import { AnnouncementExtended } from '@mauriora/model-announcement-extended';
    import { getCreateByIdOrTitle } from '@mauriora/controller-sharepoint-list';

    const newController = await getCreateByIdOrTitle(listName, siteUrl);
    const now: string = new Date().toISOString();
    const newModel = await newController.addModel(
        AnnouncementExtended,
        `(StartDate le datetime'${now}' or StartDate eq null) and (Expires ge datetime'${now}' or Expires eq null)`
    );
    if(0 === newModel.records.length ) 
    {
        await newModel.loadAllRecords();
    }
```
