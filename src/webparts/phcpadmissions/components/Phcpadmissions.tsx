import * as React from 'react';
import styles from './Phcpadmissions.module.scss';
import { IPhcpadmissionsProps } from './IPhcpadmissionsProps';
import IAdmissionItem from '../models/IAdmissionItem';
import Moment from 'react-moment';

import Carousel from 'react-multi-carousel';
import 'react-multi-carousel/lib/styles.css';


import {
	Spinner,
	SpinnerSize,
} from 'office-ui-fabric-react';


interface IPhcpadmissionsState {
  status: string;
  loading: boolean;
  items : IAdmissionItem[];
  error?: Error;
}

const responsive = {
  superLargeDesktop: {
    // the naming can be any, depends on you.
    breakpoint: { max: 4000, min: 3000 },
    items: 3,
    slidesToSlide: 3
  },
  desktop: {
    breakpoint: { max: 3000, min: 1024 },
    items: 2,
    slidesToSlide: 2
  },
  tablet: {
    breakpoint: { max: 1024, min: 600 },
    items: 2,
    slidesToSlide: 2
  },
  mobile: {
    breakpoint: { max: 600, min: 0 },
    items: 1,
    slidesToSlide: 1
  }
};

export default class Phcpadmissions extends React.Component<IPhcpadmissionsProps, IPhcpadmissionsState> {
  constructor(props: IPhcpadmissionsProps) {
    super(props);
    this.state = {
      status: this.listNotConfigured(this.props) ? 'Please configure list in Web Part properties' : 'Ready',
      loading: true,
      items: [],
    };

    if (!this.listNotConfigured(this.props)) {
      this.state = {
        status: 'Load',
        loading: false,
        items: this.props.itens,
      }
    }
    
  }

  componentDidMount(): void {
    if (!this.listNotConfigured(this.props)) {
      this.state = {
        status: 'Load',
        loading: false,
        items: this.props.itens,
      }
    }
  }

  componentDidUpdate(prevProps: Readonly<IPhcpadmissionsProps>, prevState: Readonly<IPhcpadmissionsState>, snapshot?: any): void {
    if (prevProps.itens !== this.props.itens) {
      this.setState({
        status: 'Load',
        loading: false,
        items: this.props.itens,
      });
    }
  }


  private listNotConfigured(props: IPhcpadmissionsProps): boolean {
    return props.itens === undefined ||
      props.itens === null ||
      props.itens.length === 0;
  }

  public render(): React.ReactElement<IPhcpadmissionsProps> {
    return(
      <div className={ styles.content }>
        <h2>{this.props.webparttitle}</h2>
          <div className={ styles.phcpadmissions }>        
          {this.state.status === 'Please configure list in Web Part properties' && this.state.loading ? 
            <div>
              <Spinner size={SpinnerSize.large} label={this.state.status} />
            </div>
          :
          <>
          <Carousel 
            responsive={responsive} 
            showDots={true} 
            autoPlay={true}
            autoPlaySpeed={4000}
            draggable={true}
            infinite={true}>
          {this.state.items.map((item, i) => {
              return (
                  <div key={i} className={styles.box}>
                    <div>
                      <img alt="Clock" src={require('../assets/3349548.png')} className={styles.clockImage} />
                      <Moment format="MM-DD-YYYY HH:mm">
                          {item.Created}
                      </Moment>
                    </div>
                    <p>{item.message}</p>
                    <img alt="Admissions" src={require('../assets/1453593.png')} className={styles.welcomeImage} />
                    <div className={styles.clear}></div>
                  </div>
              );
            })}
            </Carousel>
          </>  
          }
          </div>
      </div>
    );
  }  
  
}
