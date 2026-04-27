import { PrismaClient } from "@prisma/client";
import bcrypt from "bcryptjs";

const prisma = new PrismaClient();

async function main() {
  await prisma.shainmst.upsert({
    where: { shano: "dev" },
    create: {
      shano: "dev",
      shaname: "開発ユーザー",
      hshika: "1",
    },
    update: {
      shaname: "開発ユーザー",
      hshika: "1",
    },
  });

  const passwordHash = await bcrypt.hash("devpass", 10);
  await prisma.user.upsert({
    where: { employeeId: "dev" },
    create: {
      employeeId: "dev",
      passwordHash,
      role: "admin",
    },
    update: {
      passwordHash,
      role: "admin",
    },
  });
}

main()
  .then(() => prisma.$disconnect())
  .catch(async (e) => {
    console.error(e);
    await prisma.$disconnect();
    process.exit(1);
  });
