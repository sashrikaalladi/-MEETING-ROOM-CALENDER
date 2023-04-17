import { IPanelModelEnum} from '../../../controls/Event/IPanelModeEnum';
import { IEventData } from './../../../services/IEventData';
export interface ICalendarState {
  showDialog: boolean;
  eventData:IEventData ;
  selectedEvent: IEventData;
  panelMode?: IPanelModelEnum;
  startDateSlot?: any;
  endDateSlot?:Date;
  isloading: boolean;
  hasError: any;
  errorMessage: string;
  currentUserEmail:[];
  location1?:any;
}
