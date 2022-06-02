# Controller SharePoint list

Provides access to SharePoint lists. It defines the base models which can be used to extend from.

> This is part of the [hybrid repro MVC SharePoint example implementation](https://github.com/mauriora/reusable-hybrid-repo-mvc-spfx-examples)

## Table of content

- [Table of content](#table-of-content)
- [Getting Started](#getting-started)
- [Examples](#examples)
  - [Model definition](#model-definition)
    - [Built in column models](#built-in-column-models)
    - [Built in list models](#built-in-list-models)
    - [Built in base model](#built-in-base-model)
  - [Announcements](#announcements)

## Getting Started

Import this module in your project, extend a model for your list and start using it with the controller.

```shell
yarn add @mauriora/controller-sharepoint-list class-transformer reflect-metadata`
```

package.son:

```json
{
  "dependencies": {
    "@mauriora/controller-sharepoint-list": "^0.2.6",
    "class-transformer": "^0.4.0",
    "reflect-metadata": "^0.1.13"
  }
}
```

## Examples

### Model definition

Usually you would extend your model from `ListItem`. Unless a builtin model like `Announcement` exists.
Each sharepoint field is mapped to a property using the `@Expose` decorator. If the field is an object like a lookup, then it needs a `@Type` decorator. `@Exclude()` is used for properties that should be exclude from the transfer with SharePoint.

This extends the built in announcements model:

```typescript
import { Announcement, Link, UserLookup } from '@mauriora/controller-sharepoint-list';
import { Expose, Type } from 'class-transformer';

export class AnnouncementExtended extends Announcement {
    @Expose({ name: 'Urgent' })
    public urgent: boolean;

    @Expose({ name: 'StartDate' })
    public startDate: string;

    @Type( () => Link )
    @Expose({name: 'URL'})
    public url: Link;

    @Type( () => UserLookup )
    @Expose({name: 'ReportOwner'})
    public contentOwner: UserLookup;
}
```

#### Built in column models

For all complex column types a model has been created:

- Image
- Link
- MetaTerm
- User

#### Built in list models

Currently only `Announcement` is implemented.

```typescript
export class Announcement extends ListItem {
    public constructor() {
        super();
    }

    @Expose({ name: 'Body'})
    public body?: string;

    @Expose({ name: 'Expires'})
    public expires?: string;
}
```

#### Built in base model

The base model is inherited by all models. Some models extend from `DataBase` while user of this library would usually extend from `ListItem`.

The builtin base models look like follows. **The `@Exclude()` decorators have been omitted to shorten the code block**!

```typescript
export class DataBase {
    /** Source object this has been created from */
    public source: unknown;

    /** If true the item has been modified and not submitted yet*/
    public dirty = false;

    /**
     * Makes this instance observable. Needs to be called after all constructors are finished.
     * Don't call init() from inside a constructor !
     * @returns this
     */
    public init(options?: InitOptions): this {...}
}

export abstract class Deleteable extends DataBase implements IDeleteable {

    public constructor() {
        super();
    }

    abstract readonly canBeDeleted: boolean;

    abstract delete: () => Promise<void>;

    abstract deleted: boolean;
}

export class ListItemBase extends Deleteable {

    public constructor() {
        super();
    }

    public init(options?: InitOptions): this {
        ...
        return super.init(options);
    }

    @Expose({ name: 'ID' })
    public id: number;

    @Expose({ name: 'Title' })
    public title: string;

    public pnpItem: IItem | undefined;

    public deleted = false;

    public get canBeDeleted(): boolean { return (undefined !== this.pnpItem?.delete); }

    public delete = async (): Promise<void> => {...}

    public controller: SharePointList;

    public setController = (controller: SharePointList): void => {this.controller = controller; }    
}

export class ListItem extends ListItemBase {

    public constructor() {
        super();
    }

    public init(options?: InitOptions): this {
        return super.init(options);
    }

    @Type(() => UserLookup)
    @Expose({ name: 'Author', toClassOnly: true })
    author: UserLookup;

    @Expose({ name: 'Created', toClassOnly: true })
    created: string;

    @Type(() => UserLookup)
    @Expose({ name: 'Editor', toClassOnly: true })
    editor: UserLookup;

    @Expose({ name: 'Modified', toClassOnly: true })
    modified: string;

    @Expose({ name: 'Attachments' })
    public hasAttachments: boolean;

    @Expose({ name: 'ContentTypeId' })
    public contentTypeId: string;

    @Type(() => MetaTerm)
    @Expose({ name: 'TaxKeyword' })
    public taxKeyword: Array<MetaTerm> = new Array<MetaTerm>();

    @Type(() => TaxCatchAll)
    @Expose({ name: 'TaxCatchAll' })
    public taxCatchAll: Array<TaxCatchAll> = new Array<TaxCatchAll>();

    @Expose({ name: 'AverageRating', toClassOnly: true })
    public averageRating: number;

    @Expose({ name: 'RatingCount', toClassOnly: true })
    public ratingCount: number;

    @Type(() => UserLookup)
    @Expose({ name: 'RatedBy', toClassOnly: true })
    public ratedBy: Array<UserLookup> = new Array<UserLookup>();

    @Expose({ name: 'Ratings', toClassOnly: true })
    public ratings: string;    

    @Expose({ name: 'LikesCount', toClassOnly: true })
    public likesCount: number;

    @Type(() => UserLookup)
    @Expose({ name: 'LikedBy', toClassOnly: true })
    public likedBy: Array<UserLookup> = new Array<UserLookup>();

    @Expose({ name: 'EncodedAbsUrl'})
    encodedAbsUrl: string = undefined;

    public isLikedByMe = (): boolean => ...;

    public toggleLike = async (): Promise<void> => {...}

    public isRatedByMe = (): boolean=> ...;

    public myRating = (): number | undefined => {...};

    public setRating = async (rating: number): Promise<void> => {...}
}
```

### Announcements

This shows how get items from the extended Announcements List.

```typescript
    /** import the model */
    import { AnnouncementExtended } from '@mauriora/model-announcement-extended';

    /** import the controller factory */
    import { getCreateByIdOrTitle } from '@mauriora/controller-sharepoint-list';

    const newController = await getCreateByIdOrTitle(listName, siteUrl);

    const now: string = new Date().toISOString();
    /** get the SharePoint model */
    const announcements = await newController.addModel(
        AnnouncementExtended,
        `(StartDate le datetime'${now}' or StartDate eq null) and (Expires ge datetime'${now}' or Expires eq null)`
    );
    /** newModel.records is an Array of AnnouncementExtended */
    if(0 === announcements.records.length )
    {
        await announcements.loadAllRecords();
    }

    return <Stack>
        {announcements.records.map(announcement =>
            <MessageBar>
                <Text variant='large'>{announcement.title}</Text>;
            </MessageBar>
        )}
    </Stack>;
```
