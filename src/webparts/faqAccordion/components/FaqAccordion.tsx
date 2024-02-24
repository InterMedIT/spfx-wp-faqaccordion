import * as React from 'react';
import styles from './FaqAccordion.module.scss';
import { DisplayMode } from "@microsoft/sp-core-library";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import "./reactAccordion.css";

import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from "react-accessible-accordion";
import { SPFI } from "@pnp/sp";
import { getSP } from '../../../utils/pnpjs-config';
import { Web } from '@pnp/sp/webs';

export interface IFaqAccordionProps {
  webpartTitle: string;
  siteName: string;
  listName: string;
  categoryChoice: string;
  listQuestionColumn: string;
  listAnswerColumn: string;
  listSortColumn: string;
  isSortDescending: boolean;
  allowZeroExpanded: boolean;
  allowMultipleExpanded: boolean;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  onConfigure: () => void;
}

export interface IFaqAccordionState {
  items: [];
  allowMultipleExpanded: boolean;
  allowZeroExpanded: boolean;
}

export default class FaqAccordion extends React.Component<IFaqAccordionProps, IFaqAccordionState> {

  private _sp: SPFI;

  constructor(props: IFaqAccordionProps) {
    super(props);

    this.state = {
      items: [],
      allowMultipleExpanded: this.props.allowMultipleExpanded,
      allowZeroExpanded: this.props.allowZeroExpanded,
    };

    this._sp = getSP(this.context);
    this.getListItems();
  }

  private getListItems(): void {
    if (
      typeof this.props.listName !== "undefined" &&
      this.props.listName.length > 0 &&
      typeof this.props.listQuestionColumn !== "undefined" &&
      this.props.listQuestionColumn.length > 0 &&
      typeof this.props.categoryChoice !== "undefined" &&
      this.props.categoryChoice.length > 0
    ) {

      let orderByQuery = '';
      if (this.props.listSortColumn) {
        orderByQuery = `<OrderBy><FieldRef Name='${this.props.listSortColumn}' ${this.props.isSortDescending ? "Ascending='False'" : ''} /></OrderBy>`;
      }
      const query = `<View>
        <Query>
          ${orderByQuery}
          <Where>
            <Eq>
              <FieldRef Name='Category'/>
              <Value Type='Text'>${this.props.categoryChoice}</Value>
            </Eq>
          </Where>
        </Query>
      </View>`;
      const web = Web([this._sp.web, this.props.siteName]);
      const theAccordianList = web.lists.getByTitle(this.props.listName);
      theAccordianList
        .getItemsByCAMLQuery({
          ViewXml: query,
        })
        .then((results) => {
          this.setState({
            items: results,
          });
        })
        .catch((error) => {
          console.log("Failed to get list items!");
          console.log(error);
        });
    }
  }

  public componentDidUpdate(prevProps: IFaqAccordionProps): void {   
    if (
      prevProps.siteName !== this.props.siteName ||
      prevProps.listName !== this.props.listName ||
      prevProps.categoryChoice !== this.props.categoryChoice ||
      prevProps.listQuestionColumn !== this.props.listQuestionColumn ||
      prevProps.listAnswerColumn !== this.props.listAnswerColumn ||
      prevProps.isSortDescending !== this.props.isSortDescending ||
      prevProps.listSortColumn !== this.props.listSortColumn ||
      prevProps.displayMode !== this.props.displayMode
    ) {
      this.getListItems();
    }

    if (
      prevProps.allowMultipleExpanded !== this.props.allowMultipleExpanded ||
      prevProps.allowZeroExpanded !== this.props.allowZeroExpanded
    ) {
      this.setState({
        allowMultipleExpanded: this.props.allowMultipleExpanded,
        allowZeroExpanded: this.props.allowZeroExpanded,
      });
    }
  }

  public render(): React.ReactElement<IFaqAccordionProps> {
    //const { allowMultipleExpanded, allowZeroExpanded } = this.state;
    const listSelected: boolean = typeof this.props.listName !== "undefined" && this.props.listName.length > 0;
    return (
      <section className={`${styles.faqAccordion}`}>
        {!listSelected && (
          <Placeholder
            iconName="DiffInline"
            iconText="Configure your web part"
            description="Select a list with a Question, Answer and Category field to have its items rendered in a collapsible accordion format."
            buttonLabel="Choose a List"
            onConfigure={this.props.onConfigure}
          />
        )}
        {listSelected && (
          <div>
            <WebPartTitle
              displayMode={this.props.displayMode}
              title={this.props.webpartTitle}
              updateProperty={this.props.updateProperty}
            />
            <Accordion
              allowZeroExpanded={this.state.allowZeroExpanded}
              allowMultipleExpanded={this.state.allowMultipleExpanded}
            >
              {this.state.items.map((item) => {
                return (
                  <AccordionItem key={item}>
                    <AccordionItemHeading>
                      <AccordionItemButton title={item[this.props.listQuestionColumn]}>
                        {item[this.props.listQuestionColumn]}
                      </AccordionItemButton>
                    </AccordionItemHeading>
                    <AccordionItemPanel>
                      <p
                        dangerouslySetInnerHTML={{
                          __html: item[this.props.listAnswerColumn],
                        }}
                      />
                    </AccordionItemPanel>
                  </AccordionItem>
                );
              })}
            </Accordion>
          </div>
        )}
      </section>
    );
  }
}
