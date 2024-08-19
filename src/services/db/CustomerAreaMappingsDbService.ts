import { BaseDBService } from "./BaseDBService";
import config from "../../config";

export enum CustomerProcessType {
    Pod8 = "POD8",
    Transactional = "Transactional",
    Cde03 = "CDE03"
}
export interface ICustomerAreaMapping {
    name: string;
    areaPath: string;
    processType: CustomerProcessType
}

export class CustomerAreaMappingsDbService extends BaseDBService {
    public async getCustomerArea(customerName: string): Promise<ICustomerAreaMapping> {
        let ret: ICustomerAreaMapping | null = null;
        const container = await super.getDbContainer(config.cosmosDbConfig.CosmosDbContainerId_CustomerAreaMappings!);
        const querySpec = {
            query: `SELECT c.name, 
                c.areaPath, 
                t as keyword
            FROM c
            JOIN t IN c.keywords
            where CONTAINS(t,@customer,true)`,
            parameters: [
                {
                    name: '@customer',
                    value: customerName
                }
            ]
        }

        const { resources: results } = await container
            .items.query(querySpec)
            .fetchAll();
        if (results && results.length > 0) {
            ret = {
                name: results[0].name,
                areaPath: results[0].areaPath,
                processType: <CustomerProcessType>results[0].processType
            };
        }
        return ret;
    }

}