import * as React from 'react';
import styles from './Announcements.module.scss';
import { IAnnouncementsProps } from './IAnnouncementsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import pnp, { sp, Item, ItemAddResult, ItemUpdateResult } from "sp-pnp-js";
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse, HttpClientResponse } from "@microsoft/sp-http";
import * as $ from "jquery";
import ShowMoreText from 'react-show-more-text';
import { Card, Figure } from 'react-bootstrap';
export interface imyState{
  charLimit: number;
  mainDescription:string;
  listName:string;
  ID:Number;
  Remarks:String;
  arr_announcements:any[];
}
var showMore = "Read more";
var showLess = "Read less";
let _image: string = require('../Images/Announcement.png');
let _headerImage: string = require('../Images/announcements.jpg');

export default class Announcements extends React.Component<IAnnouncementsProps, imyState> {
  constructor(props: IAnnouncementsProps, state: imyState) {
    super(props);
      this.state = {
        charLimit:150,
        mainDescription:'',
        listName:this.props.description,
        ID:0,
        Remarks:'',
        arr_announcements:[]
      }
    }
    public componentDidMount(){
      $('#mainDiv').css('overflow','hidden');
      pnp.setup({
        spfxContext: this.props.currentContext
      });
      this.getLatestItemId();
      // var charLimit = this.state.charLimit;
      // var ellipsestext = "...";
      // $('.messageTag').each(function(){
      //   var content = $(this).html();
      //   if(content.length > charLimit){

      //     var c = content.substr(0, charLimit);
      //     var h = content.substr(charLimit, content.length - charLimit);
      //     var html = c + '<span class="moreellipses">' + ellipsestext+ '&nbsp;</span><span class="morecontent"><span>' + h + '</span>&nbsp;&nbsp;<a href="" class="morelink">' + showMore + '</a></span>';
 
      //     $(this).html(html);
      //   }
      // })
      // this.state.listName == this.props.listName ? this.state.listName:this.setState({listName:this.props.listName})
      // this.setState({listName:this.props.listName});
      // test = this.state.listName;
      

    }
    
  public render(): React.ReactElement<IAnnouncementsProps> {
    if($('#mainDiv').width() < 400){
      $('#mainDiv').children('div').removeClass('row');
    }
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");
    return (
      <div>
        {this.props.displayStyle == true ? <div className="row" id="header" style={{paddingLeft:'1.5rem'}}>
          {/* <img className="col-2" src={_headerImage}></img> */}
        <h3 className={styles.header + " col-10"} style={{fontSize:this.props.headerFont,color:this.props.headerColor}}>{this.props.listName}</h3>
        </div>:""}
      <div id="mainDiv" className={styles.scrollBar + " " + styles.hideScroll} onMouseEnter={this.showScroll} onMouseLeave={this.hideScroll} style={{maxHeight:'463px',padding:'.5rem'}}>
        {/* {this.state.listName == this.props.listName ? this.state.listName:this.setState({listName:this.props.listName})} */}
          {this.state.arr_announcements.map(item => {
            return (
            // <div className={styles.container + " mb-3"} style={{padding:8}}>
            <Figure className="row">
              <Figure.Image
                className="col-5"
                alt="171x180"
                src={item.Image.Url}
              />
              <Figure.Caption className={styles.description + " col-7"}>
                {item.Description}
              </Figure.Caption>
            </Figure>
          // </div>
          )
           })}
      </div>
      </div>
    );
    
  }
  private executeOnClick(isExpanded){
    console.log(isExpanded);
  }
  private showScroll(){
    $('#mainDiv').css('overflow','scroll');
    $('#mainDiv').removeClass('hideScroll');
  }
  private hideScroll(){
    $('#mainDiv').css('overflow','hidden');
  }
  private openLink(){
    window.open()
  }
  private getLatestItemId(): Promise<any> {  
    return new Promise<any>((resolve: (items: any) => void, reject: (error: any) => void): void => {  
     sp.web.lists.getByTitle(this.props.listName)  
        .items.orderBy('Order0', true).select('ID,Description,Image,Order0').get()  
        .then((items: { Description: string, Image: string, Order0: number }[]): void => {  
          if (items.length === 0) {  
            resolve("");  
          }  
          else { 
            // items.forEach(item => {
            //   this.setState({
            //     Description:item.Description
            //   })
            // }) 
            this.setState({
              arr_announcements:items
            })
            console.log(this.state.arr_announcements)
            resolve(this.state.arr_announcements);  
          }  
        }, (error: any): void => {  
          reject(error);  
        });  
    });  
  } 
}
