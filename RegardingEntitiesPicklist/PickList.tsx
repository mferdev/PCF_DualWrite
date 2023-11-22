/* eslint-disable */
import * as React from 'react';
import { IPersonaProps, IPersonaStyles } from '@fluentui/react/lib/Persona';
import {
  IBasePickerSuggestionsProps,
  PeoplePickerItemSuggestion,
  NormalPeoplePicker,
} from '@fluentui/react/lib/Pickers';
import { Service } from './service';

export interface IPickListProps {
  name?: string;
  peopleList:IPersonaProps[],
  selectedItem?:IPersonaProps[],
  pcfValue:string,
  service:Service,
  onChangePeopleList:(person: IPersonaProps[]) => void,
  onChangeSuggestion:(searchText: string) => IPersonaProps[],
  disabled:boolean
}

const suggestionProps: IBasePickerSuggestionsProps = {
  //suggestionsHeaderText: 'Suggested People',
  //mostRecentlyUsedHeaderText: 'Suggested Contacts',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading',
  showRemoveButtons: true,
  suggestionsClassName:'w-100',
  suggestionsItemClassName:'w-100',
  //suggestionsAvailableAlertText: 'People Picker Suggestions available',
  //suggestionsContainerAriaLabel: 'Suggested contacts',
  className:'w-100'
  
};

const personaStyles: Partial<IPersonaStyles> = {
  root: {
    height: 'auto',
    width:'100%'
  },
  secondaryText: {
    height: 'auto',
    whiteSpace: 'normal',
  },
  primaryText: {
    height: 'auto',
    whiteSpace: 'normal',
  },
};

export const PickList:React.FunctionComponent<IPickListProps> = (props) => {
  const picker = React.useRef(null);

  
  const onChangeItem = (item:IPersonaProps[]):void => {
    if(!item) {
      props.onChangePeopleList([]);
    }
    else props.onChangePeopleList(item);
  }

  const onRenderSuggestionItem = (personaProps: IPersonaProps, suggestionsProps: IBasePickerSuggestionsProps) => {
    return (
      <PeoplePickerItemSuggestion
        personaProps={{ ...personaProps, styles: personaStyles }}
        suggestionsProps={suggestionsProps}
      />
    );
  };
  
  function getTextFromItem(persona: IPersonaProps): string {
    return persona.text as string;
  }
  
  return (
    <div style={{width:'100%'}}>
      <NormalPeoplePicker
        // eslint-disable-next-line react/jsx-no-bind
        itemLimit={1}
        //@ts-ignore
        onResolveSuggestions={props.onChangeSuggestion}
        getTextFromItem={getTextFromItem}
        className={'ms-PeoplePicker'}
        pickerSuggestionsProps={suggestionProps}
        key={'list'}
        selectionAriaLabel={'Selected contacts'}
        removeButtonAriaLabel={'Remove'}
        selectedItems={props.selectedItem}
        //@ts-ignore
        onChange={onChangeItem}
        // eslint-disable-next-line react/jsx-no-bind
        onRenderSuggestionsItem={onRenderSuggestionItem}
        inputProps={{
          onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called ' + ev.currentTarget.value),
          onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called ' + ev.currentTarget.value),
          'aria-label': 'People Picker',
        }}
        componentRef={picker}
        resolveDelay={300}
        disabled={props.disabled}
        
      />
    </div>
  )
}




