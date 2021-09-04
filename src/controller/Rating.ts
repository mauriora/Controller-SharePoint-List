import { ListItem } from "../models/ListItem";
import {
    SPHttpClient
} from '@microsoft/sp-http';

const RATING_API = "_api/Microsoft.Office.Server.ReputationModel.Reputation.SetRating(listID=@a1,itemID=@a2,rating=@a3)";

export const setRating = async (rating: number, item: ListItem) => {
    const rateUrl = `${item.controller.site.url}/${RATING_API}?@a1='{${item.controller.listId}}'&@a2='${item.id}'&@a3=${rating}`;
    console.log(`setRating(${rating}, ${item.controller.listId}, ${item.id})`, {rateUrl, item});

    try {
        const response = await item.controller.context.spHttpClient.post( rateUrl, SPHttpClient.configurations.v1, {} );

        if(true === response.ok) {
            return;
        }
        console.error(`setRating(${rating}): Problem setting rating: ${response.statusMessage}`, {response, item})
    } catch( setRatingException ) {
        console.error(`setRating(${rating}): Problem setting rating: ${setRatingException.message}`, {setRatingException, item});
    }
}