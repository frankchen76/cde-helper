import axios, { type AxiosRequestConfig } from 'axios'
import * as https from 'https';
import { ITokenProvider } from './AzureDevOpsTokenProvider';

export enum HttpClientAuthType {
    Bearer = 'Authorization',
    ApiKey = 'api-key'
}
export class HttpClientService {
    constructor(private readonly tokenProvider: ITokenProvider, private scopes?: string[]) {

    }

    protected async getAxiosConfig(): Promise<AxiosRequestConfig> {
        const accesstoken = await this.tokenProvider.getAccessToken(this.scopes);
        //console.log("getAxiosConfig", accesstoken);
        const standardHeaders = {
            'Content-Type': 'application/json',
            cache: 'no-store'
        }
        //const headers = this.authType === HttpClientAuthType.Bearer ? { ...standardHeaders, Authorization: `Bearer ${this.accessToken}` } : { ...standardHeaders, 'x-api-key': `${this.accessToken}` }
        const headers = { ...standardHeaders, Authorization: `Bearer ${accesstoken}` }
        const config: AxiosRequestConfig = {
            headers: headers,
            httpsAgent: new https.Agent({
                rejectUnauthorized: false
            })
        }
        return config
    }

    public async get(url: string): Promise<any> {
        const config = await this.getAxiosConfig();
        const response = await axios.get(url, config)
        return response.data
    }

    public async post(url: string, body: any, isAdd?: boolean): Promise<any> {
        const config = await this.getAxiosConfig();
        const response = await axios.post(url, body, config)
        return response.data
    }

    public async patch(url: string, body: any, isAdd?: boolean): Promise<any> {
        const config = await this.getAxiosConfig();
        const response = await axios.patch(url, body, config)
        return response.data
    }

    public async delete(url: string): Promise<any> {
        const config = await this.getAxiosConfig();
        const response = await axios.delete(url, config)
        return response.data
    }
}
