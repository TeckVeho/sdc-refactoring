declare module "express-mysql-session" {
  import type session from "express-session";

  interface Options {
    host?: string;
    port?: number;
    user?: string;
    password?: string;
    database?: string;
    createDatabaseTable?: boolean;
    schema?: {
      tableName: string;
      columnNames?: {
        session_id?: string;
        expires?: string;
        data?: string;
      };
    };
  }

  type MySQLStore = typeof session.Store;

  function factory(sessionModule: typeof session): {
    new (options: Options): InstanceType<MySQLStore>;
  };

  export default factory;
}
