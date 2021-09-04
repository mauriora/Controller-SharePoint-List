import { IPrincipal } from '@pnp/spfx-controls-react';
import { Exclude, Expose } from 'class-transformer';
import { ListItemBase } from './ListItemBase';

export interface IUserFull extends Omit<IPrincipal, 'id'> {
}

export interface IUserLookup extends Pick<IPrincipal, 'title'> {

}

export class UserLookup extends ListItemBase implements IUserLookup {

    public get canBeDeleted() { return false };

    public constructor() {
        super();
    }

    @Expose({ name: 'Title' })
    public title: string = undefined;

    @Expose({ name: 'Name' })
    public claims: string = undefined;

    @Exclude()
    public picture: string = undefined;

    @Expose({ name: 'JobTitle' })
    public jobTitle: string = undefined;

    @Expose({ name: 'Department' })
    public department: string = undefined;

    @Expose({ name: 'MobilePhone'})
    public mobilePhone: string = undefined;

}

export class UserFull extends UserLookup implements IUserFull  {

    public constructor() {
        super();
    }

    @Expose({ name: 'EMail' })
    public email: string = undefined;

    @Expose({ name: 'OtherMail' })
    public otherMail: string = undefined;

    @Expose({ name: 'UserName' })
    public userName: string = undefined;

    @Expose({ name: 'UserInfoHidden' })
    public userInfoHidden: boolean = undefined

    @Expose({ name: 'ImnName'})
    public imnName: string = undefined;

    @Exclude()
    public sip: string = undefined;

    @Exclude()
    public value: string = undefined;

    // @Expose({ name: 'Id' })
    // id: string = undefined;
}
