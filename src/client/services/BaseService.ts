import { ApiKeyAuthHeader, BearerAuthHeader, HttpClientService } from "./HttpClientService";
import { IServiceContext } from "./SettingService";

export class BaseService {
    protected _httpClientService: HttpClientService;
    protected _serivceContext: IServiceContext;
    public set ServiceContext(context: IServiceContext) {
        this._serivceContext = context;
    }
    constructor() {
        this._httpClientService = new HttpClientService(new BearerAuthHeader());
    }
}