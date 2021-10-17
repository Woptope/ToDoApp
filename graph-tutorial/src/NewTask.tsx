import { useEffect, useState } from 'react';
import { NavLink as RouterNavLink, Redirect, RouteComponentProps } from 'react-router-dom';
import { Button, Col, Form, Row } from 'react-bootstrap';
import { createTask } from './GraphService';
import { useAppContext } from './AppContext';
import { Task } from './Task';

export default function NewTask(props: RouteComponentProps) {
    const app = useAppContext();

    const [taskName, setTaskName] = useState('');
    const [description, setDecription] = useState('');
    const [startDate, setStartDate] = useState('');
    const [dueDate, setDueDate] = useState('');
    const [status, setStatus] = useState('');
    const [user, setUser] = useState('');
    const [formDisabled, setFormDisabled] = useState(true);
    const [redirect, setRedirect] = useState(false);

    useEffect(() => {
        setFormDisabled(
            taskName.length === 0 ||
            startDate.length === 0 ||
            dueDate.length === 0);
    }, [taskName, startDate, dueDate]);

    const doCreate = async () => {

        const newTask: Task = {
            taskName: taskName,
            description: description,
            startDate: startDate,
            dueDate: dueDate,
            status: status

        };

        try {
            await createTask(app.authProvider!, newTask);
            setRedirect(true);
        } catch (err) {
            app.displayError!('Error creating event', JSON.stringify(err));
        }
    };

    if (redirect) {
        return <Redirect to="/todolist" />
    }

    return (
        <Form>
            <Form.Group>
                <Form.Label>Task Name</Form.Label>
                <Form.Control type="text"
                    name="taskName"
                    id="taskName"
                    className="mb-2"
                    value={taskName}
                    onChange={(ev) => setTaskName(ev.target.value)} />
            </Form.Group>
            <Form.Group as={Col} controlId="formGridState">
                <Form.Label>Status</Form.Label>
                <Form.Select
                    name="status"
                    id="status"
                    className="mb-2"
                    value={status}
                    onChange={(ev) => setStatus((ev.target as HTMLInputElement).value)}>
                    <option></option>
                    <option value="Active">Active</option>
                    <option value="Inactive">Inactive</option>
                    <option value="On Hold">On Hold</option>
                    
                    
                </Form.Select>
            </Form.Group>
            <Row className="mb-2">
                <Col>
                    <Form.Group>
                        <Form.Label>Start</Form.Label>
                        <Form.Control type="datetime-local"
                            name="startDate"
                            id="startDate"
                            value={startDate}
                            onChange={(ev) => setStartDate(ev.target.value)} />
                    </Form.Group>
                </Col>
                <Col>
                    <Form.Group>
                        <Form.Label>End</Form.Label>
                        <Form.Control type="datetime-local"
                            name="dueDate"
                            id="dueDate"
                            value={dueDate}
                            onChange={(ev) => setDueDate(ev.target.value)} />
                    </Form.Group>
                </Col>
            </Row>
            <Form.Group>
                <Form.Label>Description</Form.Label>
                <Form.Control as="textarea"
                    name="description"
                    id="description"
                    className="mb-3"
                    style={{ height: '10em' }}
                    value={description}
                    onChange={(ev) => setDecription(ev.target.value)} />
            </Form.Group>
            <Button color="primary"
                className="me-2"
                disabled={formDisabled}
                onClick={() => doCreate()}>Create</Button>
            <RouterNavLink to="/calendar"
                className="btn btn-secondary"
                exact>Cancel</RouterNavLink>
        </Form>
    );
}