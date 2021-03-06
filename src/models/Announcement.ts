import { ListItem } from './ListItem';
import { Expose } from 'class-transformer';

export class Announcement extends ListItem {
    public constructor() {
        super();
    }

    @Expose({ name: 'Body'})
    public body?: string;

    @Expose({ name: 'Expires'})
    public expires?: string;

}
