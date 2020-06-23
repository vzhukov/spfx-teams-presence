import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClient } from "@microsoft/sp-http";
import { IPresence } from '../model/IPresence';

export default class GraphService {

    private context: WebPartContext;

    constructor(private _context: WebPartContext) {
        this.context = _context;
    }

    /**
     * Gettinguser presence information
     * @param userId AAD user identity
     */
    public getPresence(userId: string): Promise<IPresence> {
        return new Promise<IPresence>((resolve, reject) => {
            this.context.msGraphClientFactory
                .getClient() // Init Microsoft Graph Client
                .then((client: MSGraphClient): Promise<IPresence> => {
                    return client
                        .api(`users/${userId}/presence`) //Get Presence method
                        .version("beta") // Beta version
                        .get((err, res) => {
                            if (err) {
                                reject(err);
                                return;
                            }
                            // Resolve presence object
                            resolve({
                                Availability: res.availability,
                                Activity: res.activity,
                            });
                        });
                });
        });
    }

    public getUserId(userUPN: string): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            this.context.msGraphClientFactory
                .getClient() // Init Microsoft Graph Client
                .then((client: MSGraphClient): Promise<IPresence> => {
                    return client
                        .api(`users/${userUPN}`) //Get Presence method
                        .version("beta") // Beta version
                        .select("id") // Select only ID attribute
                        .get((err, res) => {
                            if (err) {
                                reject(err);
                                return;
                            }
                            // Resolve presence object
                            resolve(res.id);
                        });
                });
        });
    }
}
