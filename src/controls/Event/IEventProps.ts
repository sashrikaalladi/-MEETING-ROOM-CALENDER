import { IEventData } from '../../services/IEventData';
import { IPanelModelEnum} from './IPanelModeEnum';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDateTimeFieldValue } from '@pnp/spfx-property-controls';
export interface IEventProps {
  event: IEventData;
  panelMode: IPanelModelEnum;
  onDissmissPanel: (refresh:boolean) => void;
  showPanel: boolean;
  startDate?: Date;
  endDate?: Date;
  context:WebPartContext;
  siteUrl: string;
  listId:string;
  list:any;
  eventStartDate:  IDateTimeFieldValue;
  eventEndDate: IDateTimeFieldValue;
}
