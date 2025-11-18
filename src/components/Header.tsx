import * as React from 'react';

export interface HeaderProps {
    title: string;
    logo: string;
    message: string;
    showLogout?: boolean;
    onLogout?: () => void;
}

export default class Header extends React.Component<HeaderProps> {
    render() {
        const { title, logo, message, showLogout, onLogout } = this.props;

        return (
            <section className='ms-welcome__header ms-u-fadeIn500' style={{ background: '#940427', position: 'relative' }}>
                <img src={logo} alt={title} title={title} style={{
                        height: 'auto',
                        width: '150px',
                        padding: '20px 0 5px 0',
                    }} />
                <h1 className='ms-fontSize-md ms-fontWeight-light' style={{ color: '#FFFFFF' }}>{message}</h1>
                
                {showLogout && onLogout && (
                    <button
                        onClick={onLogout}
                        title="Cerrar sesión"
                        aria-label="Cerrar sesión"
                        style={{
                            position: 'absolute',
                            top: '10px',
                            right: '10px',
                            color: '#FFFFFF',
                            backgroundColor: 'transparent',
                            border: 'none',
                            minWidth: '32px',
                            height: '32px',
                            fontSize: '18px',
                            fontWeight: 'bold',
                            cursor: 'pointer',
                            borderRadius: '4px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                        }}
                        onMouseEnter={(e) => {
                            e.currentTarget.style.backgroundColor = 'rgba(255, 255, 255, 0.1)';
                        }}
                        onMouseLeave={(e) => {
                            e.currentTarget.style.backgroundColor = 'transparent';
                        }}
                        onMouseDown={(e) => {
                            e.currentTarget.style.backgroundColor = 'rgba(255, 255, 255, 0.2)';
                        }}
                        onMouseUp={(e) => {
                            e.currentTarget.style.backgroundColor = 'rgba(255, 255, 255, 0.1)';
                        }}
                    >
                        ×
                    </button>
                )}
            </section>
        );
    }
}
