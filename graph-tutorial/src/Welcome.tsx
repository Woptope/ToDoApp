import {
    Button,
    Container
  } from 'react-bootstrap';
  import { RouteComponentProps } from 'react-router-dom';
  import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
  import { useAppContext } from './AppContext';
  
  export default function Welcome(props: RouteComponentProps) {
    const app = useAppContext();
  
    return (
      <div className="p-5 mb-4 bg-light rounded-3">
        <Container fluid>
          <h1>ToDo App</h1>
          <p className="lead">
            This is my firs React App
          </p>
          <AuthenticatedTemplate>
            <div>
              <h4>Welcome {app.user?.displayName || ''}!</h4>
              <p>Use the navigation bar at the top of the page to get started.</p>
            </div>
          </AuthenticatedTemplate>
          <UnauthenticatedTemplate>
            <Button color="primary" onClick={app.signIn!}>Click here to sign in</Button>
          </UnauthenticatedTemplate>
        </Container>
      </div>
    );
  }