import { redirect } from "next/navigation";
import { isAuthenticated } from "@/lib/auth";
import DashboardClient from "./ui";
export default async function Dashboard(){if(!(await isAuthenticated()))redirect("/login");return <DashboardClient/>}
