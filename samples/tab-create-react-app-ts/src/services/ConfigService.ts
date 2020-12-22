import * as microsoftTeams from "@microsoft/teams-js";

export interface IConfigInfo {
    shortMessage: string;
}

// Quick and dirty config stores a short string in the entity ID
export class ConfigService {

  public static getEntityId(configInfo: IConfigInfo): string {
    return configInfo.shortMessage + "/" +
      Math.random().toString(36).substring(2, 15) +
      Math.random().toString(36).substring(2, 15);
  }

  public static async getConfigInfo(): Promise<IConfigInfo> {
    return new Promise<IConfigInfo>((resolve) => {
        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            resolve({
                shortMessage: context.entityId.split('/')[0]
            });
        });
    });
  }

}