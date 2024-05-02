import * as React from 'react';
import styles from './PersonaCard.module.scss';
import { IPersonaCardProps } from './IPersonaCardProps';
import { IPersonaCardState } from './IPersonaCardState';
import { Log } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  Persona,
  PersonaSize,
  DocumentCard,
  DocumentCardType,
  Icon,
} from '@fluentui/react';

const EXP_SOURCE = 'SPFxDirectory';
const LIVE_PERSONA_COMPONENT_ID = '914330ee-2df2-4f6e-a858-30c23a812408';

export class PersonaCard extends React.Component<
  IPersonaCardProps,
  IPersonaCardState
> {
  constructor(props: IPersonaCardProps) {
    super(props);

    this.state = { livePersonaCard: undefined, pictureUrl: undefined };
  }
  /**
   *
   *
   * @memberof PersonaCard
   */
  public async componentDidMount(): Promise<void> {
    const sharedLibrary = await this._loadSPComponentById(
      LIVE_PERSONA_COMPONENT_ID
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const livePersonaCard: any = sharedLibrary.LivePersonaCard;
    this.setState({ livePersonaCard: livePersonaCard });
  }



  /**
   *
   *
   * @private
   * @returns
   * @memberof PersonaCard
   */
  private _LivePersonaCard(): JSX.Element {
    return React.createElement(
      this.state.livePersonaCard,
      {
        serviceScope: this.props.context.serviceScope,
        legacyUpn: this.props.profileProperties.Email,
        onCardOpen: () => {
          console.log('LivePersonaCard Open');
        },
        onCardClose: () => {
          console.log('LivePersonaCard Close');
        },
      },
      this._PersonaCard()
    );
  }

  /**
   *
   *
   * @private
   * @returns {JSX.Element}
   * @memberof PersonaCard
   */
  private _PersonaCard(): JSX.Element {
    return (
      <DocumentCard className={styles.documentCard} type={DocumentCardType.normal}>
        <div className={styles.profileContainer}>
          <div className={styles.profileImage}>
            <Persona
              text={this.props.profileProperties.DisplayName}
              secondaryText={this.props.profileProperties.Title}
              tertiaryText={this.props.profileProperties.Department}
              imageUrl={this.props.profileProperties.PictureUrl}
              size={PersonaSize.size72}
              imageShouldFadeIn={false}
              imageShouldStartVisible={true}
            />
            <br />
          </div>
          <div className={styles.profileInfo}>
            {/* <div className={styles.name}>{this.props.profileProperties.DisplayName}</div> */}
            {/* <div className={styles.department}>
        <Icon iconName="Gear" className={styles.icon} />
        <span>{this.props.profileProperties.Department}</span>
      </div> */}
            {this.props.profileProperties.Location && <div className={styles.location}>
              <Icon iconName="LocationDot" className={styles.icon} />
              <span>{this.props.profileProperties.Location ? this.props.profileProperties.Location : 'Not available'}</span>
            </div>}
            {this.props.profileProperties.WorkPhone && <div className={styles.phone}>
              <Icon iconName="Phone" className={styles.icon} />
              <span>{this.props.profileProperties.WorkPhone ? this.props.profileProperties.WorkPhone : 'Not available'}</span>
            </div>}
            {this.props.profileProperties.Email && <div className={styles.email}>
              <Icon iconName="Mail" className={styles.icon} />
              <span>{this.props.profileProperties.Email}</span>
            </div>}
          </div>
        </div>
      </DocumentCard>
    );
  }
  /**
   * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
   * @param componentId - componentId, guid of the component library
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async _loadSPComponentById(componentId: string): Promise<any> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const component: any = await SPComponentLoader.loadComponentById(
        componentId
      );
      return component;
    } catch (error) {
      Log.error(EXP_SOURCE, error);
      throw new Error(error);
    }
  }

  /**
   *
   *
   * @returns {React.ReactElement<IPersonaCardProps>}
   * @memberof PersonaCard
   */
  public render(): React.ReactElement<IPersonaCardProps> {
    return (
      <div className={styles.personaContainer}>
        {this.state.livePersonaCard
          ? this._LivePersonaCard()
          : this._PersonaCard()}
      </div>
    );
  }
}
