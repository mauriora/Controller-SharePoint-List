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
    public title: string;

    @Expose({ name: 'Name' })
    public claims: string;

    @Exclude()
    public picture: string;

    @Expose({ name: 'JobTitle' })
    public jobTitle: string;

    @Expose({ name: 'Department' })
    public department: string;

    @Expose({ name: 'MobilePhone'})
    public mobilePhone: string;

}

export class UserFull extends UserLookup implements IUserFull  {

    public constructor() {
        super();
    }

    @Expose({ name: 'EMail' })
    public email: string;

    @Expose({ name: 'OtherMail' })
    public otherMail: string;

    @Expose({ name: 'UserName' })
    public userName: string;

    @Expose({ name: 'UserInfoHidden' })
    public userInfoHidden: boolean = undefined

    @Expose({ name: 'ImnName'})
    public imnName: string;

    @Exclude()
    public sip: string;

    @Exclude()
    public value: string;

    // @Expose({ name: 'Id' })
    // id: string;
}
