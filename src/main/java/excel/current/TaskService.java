package excel.current;

import java.util.concurrent.*;

/**
 * Created by liubo on 2017/8/14.
 */
public class TaskService<T> {

    ExecutorService service = Executors.newFixedThreadPool(5);

    public Future<T>  doTask(Callable<T> task) throws ExecutionException, InterruptedException {
        Future<T> future = service.submit(task);
        return future;
    }
}
