import { useEffect, useState } from 'react';
import { NavLink as RouterNavLink, RouteComponentProps } from 'react-router-dom';
import { Table } from 'react-bootstrap';
import { getToDoLists, deleteTask } from './GraphService';
import { useAppContext } from './AppContext';



export default function MyTable(props: RouteComponentProps) {

    const [tasks, setTasks] = useState([]);
    const app = useAppContext();

    function Delete(Id: string) {

        deleteTask(app.authProvider!, Id)

        window.location.reload();
    }

    useEffect(() => {
        const loadTasks = async () => {
            if (app.user && tasks.length == 0) {
                try {
                    const tasks = await getToDoLists(app.authProvider!);
                    setTasks(tasks);
                } catch (err) {
                    app.displayError!((err as Error).message);
                }
            }
        };

        loadTasks();
    });

    return (
        <div>
            <RouterNavLink to="/newtask" className="btn btn-light btn-sm" exact>New task</RouterNavLink>
            <p></p>
            {tasks && <Table  table-striped  >
                <thead>
                    <tr>
                        <th>Task Name</th>
                        <th>Description</th>
                        <th>Start Date</th>
                        <th>Due Date</th>
                        <th>Status</th>
                        <th></th>
                    </tr>
                </thead>
                <tbody>
                    {tasks && tasks.map((listValue) => {
                        return (
                            <tr key={listValue['id']}>
                                <td>{listValue['fields']['TaskName2']}</td>
                                <td>{listValue['fields']['Description']}</td>
                                <td>{listValue['fields']['StartDate']}</td>
                                <td>{listValue['fields']['DueDate']}</td>
                                <td>{listValue['fields']['Status']}</td>
                                <td><button onClick={() => Delete(listValue['id'])} className="btn btn-danger">
                                    X
                                </button>

                                    <RouterNavLink to={"/edit/" + listValue['id']} className="btn btn-primary">
                                        Edit
                                    </RouterNavLink></td>
                            </tr>
                        );
                    })}
                </tbody>
            </Table>}


        </div>

    )
}