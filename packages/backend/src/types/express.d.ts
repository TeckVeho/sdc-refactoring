import type { UserRole } from "@prisma/client";

declare global {
  namespace Express {
    // eslint-disable-next-line @typescript-eslint/consistent-type-definitions
    interface User {
      id: string;
      employeeId: string;
      role: UserRole;
    }
  }
}

export {};
