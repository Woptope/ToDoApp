import NewEvent from './NewEvent';
import Calendar from './Calendar';
import MyTable from './TodoList';
import NewTask from './NewTask';
import EditTask from './EditTask';
import { BrowserRouter as Router, Route } from 'react-router-dom';
import { Container } from 'react-bootstrap';
import { MsalProvider } from '@azure/msal-react'
import { IPublicClientApplication } from '@azure/msal-browser';

import ProvideAppContext from './AppContext';
import ErrorMessage from './ErrorMessage';
import NavBar from './NavBar';
import Welcome from './Welcome';
import 'bootstrap/dist/css/bootstrap.css';
type AppProps = {
  pca: IPublicClientApplication
};

export default function App({ pca }: AppProps) {
  return(
    <MsalProvider instance={ pca }>
      <ProvideAppContext>
        <Router>
          <NavBar />
          <Container>
            <ErrorMessage />
            <Route exact path="/"
              render={(props) =>
                <Welcome {...props} />
              } />
            <Route exact path="/calendar"
              render={(props) =>
                <Calendar {...props} />
              } />
            <Route exact path="/newevent"
              render={(props) =>
                <NewEvent {...props} />
              } />
              <Route exact path="/edit/:id"
              render={(props) =>
                <EditTask {...props} />
              } />
              <Route exact path="/newtask"
              render={(props) =>
                <NewTask {...props} />
              } />
              <Route exact path="/todolist"
              render={(props) =>
                <MyTable {...props} />
              } />
          </Container>
        </Router>
      </ProvideAppContext>
    </MsalProvider>
  );
}