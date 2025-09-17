import * as restify from "restify";
import { CompletedTasksDbSerivce } from "../services/db/CompletedTasksDbSerivce";

export const TasksReportRoute = (server, passport) => {
    //server.post('/api/TaskReport', authenticateKey, async (req, res) => {
    //server.post('/api/TaskReport', authenticateBearKey, async (req: Request, res: Response) => {
    server.post('/api/TaskReport', passport.authenticate('oauth-bearer', { session: false }), async (req: restify.Request, res: restify.Response) => {
        const dbService = new CompletedTasksDbSerivce();
        try {
            //info("body", req.body);
            //const result = await dbService.addTask(req.params.upn, req.body.tasks, req.body.reportDate);
            const result = await dbService.addTask(req.query.upn, req.body.tasks, req.body.reportDate);
            //info("result", result);
            if (result && (result.statusCode == 200 || result.statusCode == 201))
                res.send(result.statusCode, { "id": result.item.id });
            else
                res.send(400, "Failed to save the task report.");
        } catch (err) {
            console.error(err);
            res.send(500, {
                error: err
            });
        }
    });
    server.get('/api/TaskReport', passport.authenticate('oauth-bearer', { session: false }), async (req: restify.Request, res: restify.Response) => {
        const dbService = new CompletedTasksDbSerivce();
        try {
            //info("upn:", req.query.upn, ";reportDate:", req.query.reportDate);
            //const result = await dbService.addTask(req.params.upn, req.body.tasks, req.body.reportDate);
            const result = await dbService.getTasks(req.query.upn, req.query.reportDate);
            //info("result", result);
            if (result && result.resources && result.resources.length > 0)
                res.send(result.resources[0]);
            else
                res.send(400, "Failed to retrieve the task report.");
        } catch (err) {
            console.error(err);
            res.send(500, {
                error: err
            });
        }
    });

}