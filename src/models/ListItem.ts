import {
    Exclude,
    Expose,
    Type
} from 'class-transformer';
import { UserLookup } from "./User";
import { ListItemBase } from './ListItemBase';
import { InitOpions } from './DataBase';
import { MetaTerm } from './MetaTerm';
import { getUserLookupSync } from '../controller/SharePoint/SharePointList';
import { TaxCatchAll } from './TaxCatchAll';
import { setRating } from '../controller/Rating';
import "@pnp/sp/comments/item";


export class ListItem extends ListItemBase {

    public constructor() {
        super();
    }

    public init(options?: InitOpions): this {
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

    @Exclude()
    public isLikedByMe = (): boolean => this.likedBy.some(
        prospect => this.controller.site.currentUser.Id === prospect.id
    );

    @Exclude()
    public toggleLike = async (): Promise<void> => {
        if (!this.pnpItem) throw new Error(`ListItem[${this.constructor.name}]#${this.id}@${this.controller?.listInfo?.Title ?? this.controller?.listId}.toggleLike(): no pnpItem, has this item been created?`)

        const mySiteUser = this.controller.site.currentUser;
        console.log(`ListItem[${this.constructor.name}]#${this.id}@${this.controller?.listInfo?.Title ?? this.controller?.listId}.toggleLike(): from=${this.isLikedByMe()}`, {item: this});

        if (this.isLikedByMe()) {
            await this.pnpItem.unlike();
            this.likedBy.splice(
                this.likedBy.findIndex(
                    prospect => mySiteUser.Id === prospect.id
                ),
                1
            );
            this.likesCount -= 1;
        } else {
            await this.pnpItem.like();
            const me = getUserLookupSync(
                mySiteUser.Id,
                this.controller.selectedFields.get('LikedBy')
            );

            this.likedBy.push( me );
            this.likesCount = (this.likesCount ?? 0) + 1;
        }
    }

    @Exclude()
    public isRatedByMe = (): boolean=> this.ratedBy.some(
        prospect => this.controller.site.currentUser.Id === prospect.id
    );

    @Exclude()
    public myRating = (): number | undefined => {
        const myUserId = this.controller.site.currentUser.Id;
        const myRatingIndex = this.ratedBy.findIndex( prospect => myUserId === prospect.id );
        if(0 <= myRatingIndex) {
            const ratingText = this.ratings.split(',')[myRatingIndex];
            return Number(ratingText);
        }
    };

    @Exclude()
    public setRating = async (rating: number): Promise<void> => {
        await setRating(rating, this);
        if(! this.isRatedByMe()) {
            const me = getUserLookupSync(
                this.controller.site.currentUser.Id,
                this.controller.selectedFields.get('RatedBy')
            );

            this.ratedBy.push( me );
            this.ratingCount = (this.ratingCount ?? 0) + 1;
            this.ratings = (this.ratings ?? '') + `${rating},`;
        } else {
            const ratings = this.ratings.split(',');
            const myUserId = this.controller.site.currentUser.Id;
            const myRatingIndex = this.ratedBy.findIndex( prospect => myUserId === prospect.id );
            ratings[myRatingIndex] = rating.toFixed();
            this.ratings = ratings.join(',');
        }
        const newRating = await this.pnpItem.select('AverageRating')();
        console.log(`ListItem[${this.constructor.name}]#${this.id}@${this.controller?.listInfo?.Title ?? this.controller?.listId}.setRating(${rating}) ${this.averageRating} => ${newRating.AverageRating}`, newRating);

        this.averageRating = newRating.AverageRating;
    }
}

export interface ListItemConstructor<ListItemType extends ListItem = ListItem> {
    new(): ListItemType;
}
