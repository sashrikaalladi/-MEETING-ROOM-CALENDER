/*eslint-disable*/ 
import * as React from 'react';
import './Calendar.module.scss';
import { ICalendarProps } from './ICalendarProps';
import { ICalendarState } from './ICalendarState';
import { debounce, escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
import * as strings from 'CalendarWebPartStrings';
import 'react-big-calendar/lib/css/react-big-calendar.css';
require('./calendar.css');
import { CommunicationColors, FluentCustomizations, FluentTheme } from '@uifabric/fluent-theme';
import Year from './Year';
import { Calendar as MyCalendar, momentLocalizer } from 'react-big-calendar';
//import Room from './Room';
import {
  Customizer,
  IPersonaSharedProps,
  Persona,
  PersonaSize,
  PersonaPresence,
  HoverCard, IHoverCard, IPlainCardProps, HoverCardType, DefaultButton,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardDetails,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  IDocumentCardPreviewImage,
  DocumentCardType,
  Label,
  ImageFit,
  IDocumentCardLogoProps,
  DocumentCardLogo,
  DocumentCardImage,
  Icon,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Dropdown,
  IDropdownOption,


} from 'office-ui-fabric-react';
import { EnvironmentType } from '@microsoft/sp-core-library';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';
import spservices from '../../../services/spservices';
import { stringIsNullOrEmpty } from '@pnp/common';
import { Event } from '../../../controls/Event/event';
import { IPanelModelEnum } from '../../../controls/Event/IPanelModeEnum';
import { IEventData } from './../../../services/IEventData';
import { IUserPermissions } from './../../../services/IUserPermissions';
import { Item } from '@pnp/pnpjs';
//import { IEventData } from './../../../services/IEventData';

//const localizer = BigCalendar.momentLocalizer(moment);
const localizer = momentLocalizer(moment);
/**
 * @export
 * @class Calendar
 * @extends {React.Component<ICalendarProps, ICalendarState>}
 */
export default class Calendar extends React.Component<ICalendarProps, ICalendarState> {
  private spService: spservices = null;
  private userListPermissions: IUserPermissions = undefined;
  private locationDropdownOption: { key: string; text: string; }[];
  public state: {
    location1:any;
    endDateSlot: any;
    startDateSlot: any;
    panelMode: any; 
    showDialog: boolean; 
    eventData:any;
   selectedEvent: any; 
   isloading: boolean; 
   hasError: boolean; 
   errorMessage: string;
    currentUserEmail:[];
  };

 public props: any;
  //public locationDropdownOption:IDropdownOption[];
  public getChoiceFieldOptions: IDropdownOption[];
  public constructor(props) {
    super(props);

    this.state = {
      //userPermissions: { hasPermissionAdd: false, hasPermissionDelete: false, hasPermissionEdit: false, hasPermissionView: false },
      location1:'',
      endDateSlot: null,
      showDialog: false,
      startDateSlot: null,  
      eventData:[],
      selectedEvent: undefined,
      isloading: true,
      hasError: false,
      panelMode:IPanelModelEnum,
      errorMessage: '',
      currentUserEmail:[],
      //userPermissions: { hasPermissionAdd: false, hasPermissionDelete: false, hasPermissionEdit: false, hasPermissionView: false },
    };
    this.onLocationChanged = this.onLocationChanged.bind(this);
    this.onDismissPanel = this.onDismissPanel.bind(this);
    this.onSelectEvent = this.onSelectEvent.bind(this);
    this.onSelectSlot = this.onSelectSlot.bind(this);
    this.spService = new spservices(this.props.context);
    moment.locale(this.props.context.pageContext.cultureInfo.currentUICultureName);

  }

  private onDocumentCardClick(ev: React.SyntheticEvent<HTMLElement, Event>) {
    ev.preventDefault();
    ev.stopPropagation();
  }
  /**
   * @private
   * @param {*} event
   * @memberof Calendar
   */
  //  setState(_arg0: {isloading:boolean; showDialog: boolean; selectedEvent: any; panelMode: IPanelModelEnum; }): void {
  //    throw new Error('Method not implemented.');
  // }
  private async onSelectEvent(event: any,{start,end}) {
    
    //console.log("hai this is on event click",this.context.user,event.ownerName,event,this.props.siteUrl);
    try{
    const currentUser:any= await this.spService.getCurrentUser(this.props.siteUrl);
    console.log(currentUser);
   let userobj= currentUser.Email;
  
    console.log(userobj,event.ownerEmail,event.ownerName);
    if(userobj===event.ownerEmail){
      this.setState({ isloading: false, showDialog: true, selectedEvent: event, panelMode: IPanelModelEnum.edit });
      }
      else{
        this.setState({ showDialog: false, startDateSlot: start, endDateSlot: end, selectedEvent: undefined, panelMode: IPanelModelEnum.view});
      }
//this.setState({ isloading: false, showDialog: true, selectedEvent: event, panelMode: IPanelModelEnum.edit });
    }
    catch(e){
      console.log(e)
    }
    
   
  }


  /**
   *
   * @private
   * @param {boolean} refresh
   * @memberof Calendar
   */
  private async onDismissPanel(refresh: boolean): Promise<void> {

    this.setState({
      showDialog: false,
      selectedEvent: undefined,
      isloading: false,
      panelMode: IPanelModelEnum.add
    });
    if (refresh === true) {
      this.setState({
        isloading: true,
        showDialog: false,
        selectedEvent: undefined,
        panelMode: IPanelModelEnum.add
      });
      await this.loadEvents(this.state.location1);
      this.setState({
        isloading: false,
        showDialog: false,
        selectedEvent: undefined,
        panelMode: IPanelModelEnum.add
      });
    }
  }
  /**
   * @private
   * @memberof Calendar
   */
  private async loadEvents(resourceRoom:any) {
    try {
      // Teste Properties
      if (!this.props.list || !this.props.siteUrl || !this.props.eventStartDate.value || !this.props.eventEndDate.value) return;

      this.userListPermissions = await this.spService.getUserPermissions(this.props.siteUrl, this.props.list);
      console.log(this.props.siteUrl,this.props.list,this.props.eventStartDate.value,this.props.eventEndDate.value);
    
      if(!resourceRoom){
      const eventsData:any = await this.spService.getEvents(escape(this.props.siteUrl), escape(this.props.list), this.props.eventStartDate.value, this.props.eventEndDate.value);
      this.setState({ eventData: eventsData, hasError: false, errorMessage: "" }); 
    }
      else{
        console.log(resourceRoom);
        try{
        const eventsData:any = await this.spService.getEventsByResourceRoom(escape(this.props.siteUrl), escape(this.props.list), this.props.eventStartDate.value, this.props.eventEndDate.value,resourceRoom);
        // console.log(eventsData,this.props.eventEndDate.value);
        this.setState({ eventData: eventsData, hasError: false, errorMessage: "" });
        }
        catch(e){
          console.log(e);
        }
      }
     
    } catch (error) {
      console.log(error);
      this.setState({ hasError: true, errorMessage: error.message, isloading: false });
    }
  }
  /**
   * @memberof Calendar
   */
  public async componentDidMount() {
    
    this.setState({ isloading: true });
    await this.loadEvents(this.state.location1);
    this.setState({ isloading: false});
    if(!this.state.location1)
    {
    await this.renderEventData();
    }
     this.setState({ isloading: false });
   }

  /**
   *
   * @param {*} error
   * @param {*} errorInfo
   * @memberof Calendar
   */
  public componentDidCatch(error: any, errorInfo: any): void {
    this.setState({ hasError: true, errorMessage: errorInfo.componentStack });
  }
  /**
   *
   *
   * 
   * @param {ICalendarProps} prevProps
   * @param {ICalendarState} prevState
   * @memberof Calendar
   */
  public async componentDidUpdate(prevProps: ICalendarProps, prevState: ICalendarState) {

    if (!this.props.list || !this.props.siteUrl || !this.props.eventStartDate.value || !this.props.eventEndDate.value) return;
    // Get  Properties change
    if (prevProps.list !== this.props.list || this.props.eventStartDate.value !== prevProps.eventStartDate.value || this.props.eventEndDate.value !== prevProps.eventEndDate.value) {
      this.setState({
        isloading: true,
        showDialog: false,
        selectedEvent: undefined,
        panelMode: IPanelModelEnum.add
      });
      await this.loadEvents(this.state.location1);
      this.setState({
        isloading: false,
        showDialog: false,
        selectedEvent: undefined,
        panelMode: IPanelModelEnum.add
      });
    }
  }
  /**
   * @private
   * @param {*} { event }
   * @returns
   * @memberof Calendar
   */
  private renderEvent({ event }) {

    const previewEventIcon: IDocumentCardPreviewProps = {
      previewImages: [
        {
          // previewImageSrc: event.ownerPhoto,
          previewIconProps: { iconName: event.fRecurrence === '0' ? 'Calendar' : 'RecurringEvent', styles: { root: { color: event.color } }, className: "previewEventIcon " },
          height: 43,
        }
      ]
    };
    const EventInfo: IPersonaSharedProps = {
      imageInitials: event.ownerInitial,
      imageUrl: event.ownerPhoto,
      text: event.title
    };

    /**
     * @returns {JSX.Element}
     */
    const onRenderPlainCard = (): JSX.Element => {
      return (
        <div className={"plainCard"}>
          <DocumentCard className={"Documentcard"}   >
            <div>
              <DocumentCardPreview {...previewEventIcon} />
            </div>
            <DocumentCardDetails>
              <div className={"DocumentCardDetails"}>
                <DocumentCardTitle title={event.title} shouldTruncate={true} className={"DocumentCardTitle"} styles={{ root: { color: event.color } }} />
              </div>

              <div className={"DocumentCardDetails"}>
                <DocumentCardTitle title={event.location1} shouldTruncate={true} className={"DocumentCardTitle"} styles={{ root: { color: event.color } }} />
              </div>

              {
                moment(event.EventDate).format('YYYY/MM/DD') !== moment(event.EndDate).format('YYYY/MM/DD') ?
                  <span className={"DocumentCardTitleTime"}>{moment(event.EventDate).format('dddd')} - {moment(event.EndDate).format('dddd')} </span> :
                  //<span className={"DocumentCardTitleTime"}>{moment(event.EventDate).format('HH:mm')}H - {moment(event.EndDate).format('HH:mm')}H</span>:
                  <span className={"DocumentCardTitleTime"}>{moment(event.EventDate).format('dddd')} </span>
              }
              {<span className={"DocumentCardTitleTime"}>{moment(event.EventDate).format('HH:mm')} - {moment(event.EndDate).format('HH:mm')}</span>}
              {/* <span className={"DocumentCardTitleTime"}>{moment(event.EventDate).format('HH:mm')}H - {moment(event.EndDate).format('HH:mm')}H</span>
              <Icon iconName='MapPin' className={"locationIcon"} style={{ color: event.color }} /> */}
              {/* <DocumentCardTitle
                title={`${event.location}`}
                shouldTruncate={true}
                showAsSecondaryTitle={true}
                className={"location"}
              /> */}
              <div style={{ marginTop: 20 }}>
                <DocumentCardActivity
                  activity={strings.EventOwnerLabel}
                  people={[{ name: event.ownerName, profileImageSrc: event.ownerPhoto, initialsColor: event.color }]}
                />
              </div>
            </DocumentCardDetails>
          </DocumentCard>
        </div>
      );
    };

    return (
      <div style={{ height: 22 }}>
        <HoverCard
          cardDismissDelay={1000}
          type={HoverCardType.plain}
          plainCardProps={{ onRenderPlainCard: onRenderPlainCard }}
          onCardHide={(): void => {
          }}
        >
          <Persona
            {...EventInfo}
            size={PersonaSize.size24}
            presence={PersonaPresence.none}
            coinSize={22}
            initialsColor={event.color}
          />
        </HoverCard>
      </div>
    );
  }
  /**
   *
   *
   * @private
   * @memberof Calendar
   */
  private onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  /**
   * @param {*} { start, end }
   * @memberof Calendar
   */
  public async onSelectSlot({ start, end }) {
    if (!this.userListPermissions.hasPermissionAdd) return;
    this.setState({ showDialog: true, startDateSlot: start, endDateSlot: end, selectedEvent: undefined, panelMode: IPanelModelEnum.add });
    
  }

  /**
   *
   * @param {*} event
   * @param {*} start
   * @param {*} end
   * @param {*} isSelected
   * @returns {*}
   * @memberof Calendar
   */
  public eventStyleGetter(event, start, end, isSelected): any {

    let style: any = {
      backgroundColor: 'white',
      borderRadius: '0px',
      opacity: 1,
      color: event.color,
      borderWidth: '1.1px',
      borderStyle: 'solid',
      borderColor: event.color,
      borderLeftWidth: '6px',
      display: 'block'
    };

    return {
      style: style
    };
  }


  /**
    *
    * @param {*} date
    * @memberof Calendar
    */
  public dayPropGetter(date: Date) {
    return {
      className: "dayPropGetter"
    };
  }

  /**
     *
     *
     * @private
     * @memberof Event
     */
debugger
  private async renderEventData() {
    console.log("checkfunction");
    //this.setState({ isloading: true });
    try {
      console.log("tried");
      console.log(this.props.list, this.props.listId);
      // const userListPermissions: IUserPermissions = await this.spService.getUserPermissions(this.props.siteUrl, this.props.listId);
      this.locationDropdownOption = await this.spService.getChoiceFieldOptions(this.props.siteUrl, this.props.list, 'location1');//to give field name of the list
      console.log(this.locationDropdownOption);
    }
    catch (e) {
      console.log(e);
    }
  }



  /**
  *
  * @private
  * @param {React.FormEvent<HTMLDivElement>} ev
  * @param {IDropdownOption} item
  * @memberof Event
  * @param item
  */
  public async onLocationChanged(ev: React.FormEvent<HTMLDivElement>, item: IDropdownOption):Promise<any> {
   this.setState({...this.state.eventData});
   this.setState({location1:item.text})
   //let x = this.state.location1;
   try{
   this.setState({isloading:true});
   await this.loadEvents(item.text);
    this.setState({isloading:false});
    }
    catch(error){
      console.log(error);
    }
    
    
  }

  /**
   *
   * @returns {React.ReactElement<ICalendarProps>}
   * @memberof Calendar
   */
  public render(): React.ReactElement<ICalendarProps> {
    return (
      <>
       
     
        <Customizer>


          <div className={"calendar"} style={{ backgroundColor: 'white', padding: '20px'}}>
            <h2>Meeting Room Reservation</h2>
              <div style={{padding: '20px'}}>
           <Dropdown
            label={strings.locationLabel}
           options={this.locationDropdownOption}
          selectedKey={this.state.eventData  && this.state.location1? this.state.location1:''}
           //selectedKey={this.state.location1}
          onChange={this.onLocationChanged}
            placeholder={strings.locationPlaceHolder}
          
          //disabled={this.state.userPermissions.hasPermissionAdd || this.state.userPermissions.hasPermissionEdit ? false : true}
          />
        </div>
            {
              (!this.props.list || !this.props.eventStartDate.value || !this.props.eventEndDate.value) ?
                <Placeholder iconName='Edit'
                  iconText={strings.WebpartConfigIconText}
                  description={strings.WebpartConfigDescription}
                  buttonLabel={strings.WebPartConfigButtonLabel}
                  hideButton={this.props.displayMode === DisplayMode.Read}
                  onConfigure={this.onConfigure.bind(this)} />
                :
                // test if has errors
                this.state.hasError ?
                  <MessageBar messageBarType={MessageBarType.error}>
                    {this.state.errorMessage}
                  </MessageBar>
                  :
                  // show Calendar
                  // Test if is loading Events
                  <div>
                    {this.state.isloading ? <Spinner size={SpinnerSize.large} label={strings.LoadingEventsLabel} /> :
                      <div className={"container"}>
                        <MyCalendar
                          dayPropGetter={this.dayPropGetter}
                          localizer={localizer}
                          selectable
                          events={this.state.eventData}
                          startAccessor="EventDate"
                          endAccessor="EndDate"
                          eventPropGetter={this.eventStyleGetter}
                          onSelectSlot={this.onSelectSlot}
                          components={{
                            event: this.renderEvent
                          }}
                          onSelectEvent={this.onSelectEvent}
                          defaultDate={moment().startOf('day').toDate()}
                          views={{
                            day: true,
                            week: true,
                            month: true,
                            // agenda: true,
                            work_week: Year
                          }}
                          messages={
                            {
                              'today': strings.todayLabel,
                              'previous': strings.previousLabel,
                              'next': strings.nextLabel,
                              'month': strings.monthLabel,
                              'week': strings.weekLabel,
                              'day': strings.dayLable,
                              'room': strings.RoomLable,
                              'showMore': total => `+${total} ${strings.showMore}`,
                              'work_week': strings.yearHeaderLabel
                            }
                          }
                        />
                      </div>
                      
                    }
                  </div>
            }
           {
              
            
              this.state.showDialog &&
              <Event
                event={this.state.selectedEvent}
                panelMode={this.state.panelMode}
                onDissmissPanel={this.onDismissPanel}
                showPanel={this.state.showDialog}
                startDate={this.state.startDateSlot}
                endDate={this.state.endDateSlot}
                context={this.props.context}
                siteUrl={this.props.siteUrl}
                listId={this.props.list}
                list={this.props.list}
                eventStartDate={this.props.eventStartDate}
                 eventEndDate={this.props.eventEndDate}
              />
              
            }

          </div>

         

        </Customizer>
        
        
      </>
    );
  }
}
