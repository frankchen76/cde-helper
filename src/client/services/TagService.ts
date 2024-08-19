import { BaseService } from "./BaseService"
import { ISettingItem } from "./SettingService";
import { TagCollection } from "./Tag";

export interface ITagService {
    getTags(settingItem: ISettingItem): Promise<TagCollection>;
}
export class TagService extends BaseService implements ITagService {
    public async getTags(settingItem: ISettingItem): Promise<TagCollection> {
        let ret = null;
        let url = `${settingItem.baseUrl}/_apis/wit/tags?api-version=6.0-preview.1`;
        const response = await this._httpClientService.get(url);
        if (response != null) {
            ret = TagCollection.createTagsFromResponse(response);
        }
        return ret;
    }

}