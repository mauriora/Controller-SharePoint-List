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


export class ListItem extends ListItemBase {

    public constructor() {
        super();
    }

    public init(options?: InitOpions) {
        return super.init(options);
    }

    @Type(() => UserLookup)
    @Expose({ name: 'Author', toClassOnly: true })
    author: UserLookup = undefined;

    @Expose({ name: 'Created', toClassOnly: true })
    created: string = undefined;

    @Type(() => UserLookup)
    @Expose({ name: 'Editor', toClassOnly: true })
    editor: UserLookup = undefined;

    @Expose({ name: 'Modified', toClassOnly: true })
    modified: string = undefined;

    @Expose({ name: 'Attachments' })
    public hasAttachments: boolean = undefined;

    @Expose({ name: 'ContentTypeId' })
    public contentTypeId: string = undefined;

    @Type(() => MetaTerm)
    @Expose({ name: 'TaxKeyword' })
    public taxKeyword: Array<MetaTerm> = new Array<MetaTerm>();

    @Type(() => TaxCatchAll)
    @Expose({ name: 'TaxCatchAll' })
    public taxCatchAll: Array<TaxCatchAll> = new Array<TaxCatchAll>();

    @Expose({ name: 'AverageRating', toClassOnly: true })
    public averageRating: number = undefined;

    @Expose({ name: 'RatingCount', toClassOnly: true })
    public ratingCount: number = undefined;

    @Type(() => UserLookup)
    @Expose({ name: 'RatedBy', toClassOnly: true })
    public ratedBy: Array<UserLookup> = new Array<UserLookup>();

    @Expose({ name: 'Ratings', toClassOnly: true })
    public ratings: string = undefined;    

    @Expose({ name: 'LikesCount', toClassOnly: true })
    public likesCount: number = undefined;

    @Type(() => UserLookup)
    @Expose({ name: 'LikedBy', toClassOnly: true })
    public likedBy: Array<UserLookup> = new Array<UserLookup>();

    @Exclude()
    public isLikedByMe = () => this.likedBy.some(
        prospect => this.controller.site.currentUser.Id === prospect.id
    );

    @Exclude()
    public toggleLike = async () => {
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
            this.likesCount += 1;
        }
    }

    @Exclude()
    public isRatedByMe = () => this.ratedBy.some(
        prospect => this.controller.site.currentUser.Id === prospect.id
    );

    @Exclude()
    public myRating = () => {
        console.log(`ListItem[${this.constructor.name}]#${this.id}@${this.controller?.listInfo?.Title ?? this.controller?.listId}.myRating(${this.controller?.site?.currentUser?.Id})`, {ratings: this.ratings, ratedBy: this.ratedBy});

        const myUserId = this.controller.site.currentUser.Id;
        const myRatingIndex = this.ratedBy.findIndex( prospect => myUserId === prospect.id );
        if(0 <= myRatingIndex) {
            const ratingText = this.ratings.split(',')[myRatingIndex];
            console.log(`ListItem[${this.constructor.name}]#${this.id}@${this.controller?.listInfo?.Title ?? this.controller?.listId}.myRating(${this.controller?.site?.currentUser?.Id}) = ${ratingText}`, {ratings: this.ratings, ratedBy: this.ratedBy});
            return Number(ratingText);
        }
    };

    @Exclude()
    public setRating = async (rating: number) => {
        await setRating(rating, this);
        if(! this.isRatedByMe()) {
            const me = getUserLookupSync(
                this.controller.site.currentUser.Id,
                this.controller.selectedFields.get('RatedBy')
            );

            this.ratedBy.push( me );
            this.ratingCount += 1;
            this.ratings += `${rating},`;
        } else {
            const ratings = this.ratings.split(',');
            const myUserId = this.controller.site.currentUser.Id;
            const myRatingIndex = this.ratedBy.findIndex( prospect => myUserId === prospect.id );
            ratings[myRatingIndex] = rating.toFixed();
            this.ratings = ratings.join(',');
        }
        const newRating = await this.pnpItem.select('AverageRating').get();
        console.log(`ListItem[${this.constructor.name}]#${this.id}@${this.controller?.listInfo?.Title ?? this.controller?.listId}.setRating(${rating}) ${this.averageRating} => ${newRating.AverageRating}`, newRating);

        this.averageRating = newRating.AverageRating;
    }
}

export interface ListItemConstructor<ListItemType extends ListItem = ListItem> {
    new(): ListItemType;
}
