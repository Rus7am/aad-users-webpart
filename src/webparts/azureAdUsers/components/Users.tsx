import * as React from 'react';
import styles from './Users.module.scss';

import { User } from '@microsoft/microsoft-graph-types';
import { Persona, IPersonaSharedProps, PersonaInitialsColor, PersonaSize } from 'office-ui-fabric-react';

import { IUsersProps } from './IUsersProps';

import { AzureADService } from '../../../services/AzureADService';

export default class Users extends React.Component<IUsersProps, {
  users: User[],
  error: boolean
}> {
  constructor(props) {
    super(props);

    this.state = {
      users: [],
      error: false
    };
  }

  public componentDidMount() {
    AzureADService.getUsers().then((users) => this.setState({
      users: users,
      error: false
    })).catch((error: Error) => {
      this.setState({
        error: true
      });
      console.error(error.message);
    });
  }

  public render(): React.ReactElement<IUsersProps> {
    const { users, error } = this.state;

    if (users.length === 0 && !error) {
      return (
        <div className={styles.users}>
          <div>
            Loading users
          </div>
        </div>
      );
    }
    if (error) {
      return (
        <div className={styles.users}>
          <div className={styles.error}>
            Error loading users
          </div>
        </div>
      );
    }
    return (
      <div className={styles.users}>
        <div className={styles.container}>
          <div className={styles.row}>
            {users.map((user, index) =>
              <div className={styles.column} key={index}>
                <Persona initialsColor={PersonaInitialsColor.teal} size={PersonaSize.size72}
                  text={user.displayName} secondaryText={user.mail} tertiaryText={user.businessPhones[0]} />
              </div>)}
          </div>
        </div>
      </div>
    );
  }
}
