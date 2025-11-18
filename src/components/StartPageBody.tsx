import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import HeroList, { HeroListItem } from './HeroList';

export interface StartPageBodyProps {
    listItems: HeroListItem[];
    login: () => {};
}

export default class StartPageBody extends React.Component<StartPageBodyProps> {
    render() {
        const { listItems, login } = this.props;

        return (
            <div className='ms-welcome'>
                <HeroList message='Por favor refresque su inicio de sesión' items={listItems}>
                </HeroList>
                <div className='ms-welcome__main'>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={login}>Conectarse a Office 365</Button>
                </div>
            </div>
        );
    }
}
