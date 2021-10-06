import { Expose, Type } from "class-transformer";
import { DataBase } from "./DataBase";

// {
//     "fileName": "Set-10791.png",
//     "serverRelativeUrl": "/sites/ResourcesDev/SiteAssets/Lists/13341d9d-6f81-47c0-a207-116a35eb7fe6/Set-107912.png",
//     "id": "a695d1ef-4c03-4393-8f96-9e97809877be",
//     "serverUrl": "https://pito.sharepoint.com",
//     "thumbnailRenderer": { 
//         "spItemUrl": "https://pito.sharepoint.com:443/_api/v2.1/drives/b!_YaK12LmeEqY3t8_2uRx0v_0JouCVt1PvWl6n875EO0RPvV2HACWSLLd7aZMTohZ/items/015PAA4IPP2GK2MA2MSNBY7FU6S6AJQ556",
//         "fileVersion": 1,
//         "sponsorToken": "L3NpdGVzL1Jlc291cmNlc0Rldi9MaXN0cy9SZXNvdXJjZVNldHxQcmV2aWV3fDU"
//     },
//     "type": "thumbnail",
//     "fieldName": "Preview"
// }

export class ThumbnailRenderer extends DataBase {
    constructor() {
        super();
    }

    @Expose({ name: 'spItemUrl' })
    public spItemUrl: string;

    @Expose({ name: 'fileVersion' })
    public fileVersion: number;

    @Expose({ name: 'sponsorToken' })
    public sponsorToken: string;

    public static is = (prospect: unknown): prospect is ThumbnailRenderer => {
        return ('string' === typeof (prospect as ThumbnailRenderer).spItemUrl) &&
            ('string' === typeof (prospect as ThumbnailRenderer).fileVersion) &&
            ('string' === typeof (prospect as ThumbnailRenderer).sponsorToken);
    }
}

export class Image extends DataBase {
    constructor() {
        super();
    }

    @Expose({ name: 'fileName' })
    public fileName: string;

    @Expose({ name: 'serverRelativeUrl' })
    public serverRelativeUrl: string;

    @Expose({ name: 'serverUrl' })
    public serverUrl: string;

    @Expose({ name: 'type' })
    public type: string;

    @Expose({ name: 'fieldName' })
    public fieldName: string;

    @Type( () => ThumbnailRenderer )
    @Expose({ name: 'thumbnailRenderer'})
    public thumbnailRenderer: ThumbnailRenderer;

    public static is = (prospect: unknown): prospect is Image => {
        return ('string' === typeof (prospect as Image).fileName) &&
            ('string' === typeof (prospect as Image).serverRelativeUrl) &&
            ('string' === typeof (prospect as Image).serverUrl) &&
            ('string' === typeof (prospect as Image).type) &&
            ('string' === typeof (prospect as Image).fieldName) &&
            ('object' === typeof (prospect as Image).thumbnailRenderer) &&
            ThumbnailRenderer.is((prospect as Image).thumbnailRenderer);
    }
}
