import bcrypt from "bcryptjs";
import readline from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";

const rl = readline.createInterface({ input, output });
const password = await rl.question("Senha do administrador: ");
rl.close();
console.log(await bcrypt.hash(password, 12));
