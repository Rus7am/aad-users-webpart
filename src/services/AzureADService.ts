import { graph } from '@pnp/graph';
import { User } from '@microsoft/microsoft-graph-types';

// TODO: Implement scoped service, see https://www.vrdmn.com/2019/03/using-service-scopes-to-decouple.html
export class AzureADService {
  public static getUsers(top: number = 10): Promise<User[]> {
    return graph.users.top(top).get();
  }
}
