import { BaseService } from "./BaseService";
import { QueryCollection } from "./Query";

export interface IQueryService {
    getQueries(baseUrl: string): Promise<QueryCollection>;
}
export class QueryService extends BaseService {

    public async getQueries(baseUrl: string): Promise<QueryCollection> {
        let url = `${baseUrl}/_apis/wit/queries/?api-version=6.0&$depth=1`;
        const queryResponse = await this._httpClientService.get(url);
        const ret = QueryCollection.createQueriesFromResponse(queryResponse);

        return ret;
    }

}