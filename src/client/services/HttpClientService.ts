import { ClientSideAuthService } from "./ClientSideAuthService";

interface IAuthHeader {
    getHeaders(isAdd: boolean): Promise<HeadersInit>;
}
export class BearerAuthHeader implements IAuthHeader {
    private _authService = new ClientSideAuthService();
    public async getHeaders(isAdd: boolean = false): Promise<HeadersInit> {
        const token = await this._authService.getAccessToken();
        return {
            'Content-Type': isAdd ? 'application/json-patch+json' : 'application/json',
            'Authorization': `Bearer ${token.access_token}`,
            'cache': "no-store"
        }
    }
}
export class ApiKeyAuthHeader implements IAuthHeader {
    constructor(private _apiKey: string) {

    }
    public async getHeaders(isAdd: boolean = false): Promise<HeadersInit> {
        return {
            'Content-Type': 'application/json',
            'x-api-key': `${this._apiKey}`
        }
    }

}
export class HttpClientService {
    private _authService = new ClientSideAuthService();
    constructor(private _authHeader: IAuthHeader = new BearerAuthHeader()) {

    }
    public async get(url: string): Promise<any> {
        const token = await this._authService.getAccessToken();
        const header = await this._authHeader.getHeaders(false);
        const response = await fetch(url, {
            method: 'GET',
            headers: header
            // headers: {
            //     'Content-Type': 'application/json',
            //     'Authorization': `Bearer ${token.access_token}`,
            //     'cache': "no-store"
            // }
        });
        if (!response.ok) {
            throw await response.text();
        }
        return response.json();
    }
    public async post(url: string, body: any, isAdd?: boolean): Promise<any> {
        const token = await this._authService.getAccessToken();
        const header = await this._authHeader.getHeaders(isAdd);
        const response = await fetch(url, {
            method: 'POST',
            headers: header,
            // headers: {
            //     'Content-Type': isAdd ? 'application/json-patch+json' : 'application/json',
            //     'Authorization': `Bearer ${token.access_token}`,
            //     'cache': "no-store"
            // },
            body: JSON.stringify(body)
        });
        if (!response.ok) {
            throw await response.text();
        }
        return response.json();
        // return fetch(url, {
        //     method: 'POST',
        //     headers: {
        //         'Content-Type': isAdd ? 'application/json-patch+json' : 'application/json',
        //         'Authorization': `Bearer ${token.access_token}`,
        //         'cache': "no-store"
        //     },
        //     body: body
        // }).then(response => response.json());
    }
    public async patch(url: string, body: any, isAdd?: boolean): Promise<any> {
        const token = await this._authService.getAccessToken();
        const header = await this._authHeader.getHeaders(isAdd);
        const response = await fetch(url, {
            method: 'PATCH',
            headers: header,
            // headers: {
            //     'Content-Type': isAdd ? 'application/json-patch+json' : 'application/json',
            //     'Authorization': `Bearer ${token.access_token}`,
            //     'cache': "no-store"
            // },
            body: JSON.stringify(body)
        });
        if (!response.ok) {
            throw await response.text();
        }
        return response.json();
        // return fetch(url, {
        //     method: 'PATCH',
        //     headers: {
        //         'Content-Type': isAdd ? 'application/json-patch+json' : 'application/json',
        //         'Authorization': `Bearer ${token.access_token}`,
        //         'cache': "no-store"
        //     },
        //     body: body
        // }).then(response => response.json());
    }
    public async delete(url: string): Promise<any> {
        const token = await this._authService.getAccessToken();
        const header = await this._authHeader.getHeaders(false);

        return fetch(url, {
            method: 'DELETE',
            headers: header
            // headers: {
            //     'Content-Type': 'application/json',
            //     'Authorization': `Bearer ${token.access_token}`,
            //     'cache': "no-store"
            // }
        }).then(response => response.json());
    }
}