import { HttpClientService } from "./HttpClientService";

export class ProfileService {
    protected _httpClientService: HttpClientService;
    constructor() {
        this._httpClientService = new HttpClientService();
    }
    public static async getUpn(): Promise<string> {
        let ret: string = null;
        const httpCliet = new HttpClientService();
        const url = "https://app.vssps.visualstudio.com/_apis/profile/profiles/me?api-version=7.1-preview.3";
        const response = await httpCliet.get(url);
        if (response && response["emailAddress"]) {
            ret = response["emailAddress"];
        }
        return ret;

    }
}