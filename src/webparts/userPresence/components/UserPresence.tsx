import * as React from 'react';
import styles from './UserPresence.module.scss';
import { IUserPresenceProps, IUserPresenceState } from './IUserPresenceProps';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Button } from 'office-ui-fabric-react';
import GraphService from '../../../services/GraphService';

export default class UserPresence extends React.Component<IUserPresenceProps, IUserPresenceState> {
  constructor(props: IUserPresenceProps) {
    super(props);
    this.state = {};

    // Bind Handlers
    this._buttonClick = this._buttonClick.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);

    // Init Microsoft Graph Service
    this._graphService = new GraphService(this.props.context);
  }

  private _graphService: GraphService

  public render(): React.ReactElement<IUserPresenceProps> {
    return (
      <div className={styles.userPresence}>
        {/* People Picker */}
        <div className={styles.container}>
          <PeoplePicker context={this.props.context} titleText="Step 1. Find user"
            selectedItems={this._getPeoplePickerItems} principalTypes={[PrincipalType.User]} />
        </div>
        {/* / People Picker */}

        {/* Button */}
        <div className={styles.container}>
          <Button text="Step 2. Get presence" className={styles.button}
            disabled={!this.state.userUPN} onClick={this._buttonClick} />
        </div>
        {/* / Button */}

        {/* State */}
        <div className={styles.container}>
          <label>useUPN:</label> {this.state.userUPN}<br />
          <label>userId:</label> {this.state.userId}<br />
          <label>Presence.Activity:</label> {this.state.presence ? this.state.presence.Activity : ""}<br />
          <label>Presence.Availability:</label> {this.state.presence ? this.state.presence.Availability : ""}
        </div>
        {/* / State */}
      </div>
    );
  }

  private _buttonClick() {
    this._graphService.getUserId(this.state.userUPN)
      .then(userId => {

        // Update User Id
        this.setState({
          userId: userId
        });

        this._graphService.getPresence(this.state.userId)
          .then(presence => {

            // Update Presence
            this.setState({ presence: presence });
          });
      });
  }

  private _getPeoplePickerItems(items: any[]): void {

    // Break if number of users does not equal 1
    if (items.length != 1) {
      return;
    }

    // Update User UPN
    this.setState({
      userUPN: items[0].secondaryText
    });


    console.log('Items:', items);
  }
}
