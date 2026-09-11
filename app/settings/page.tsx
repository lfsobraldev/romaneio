import { redirect } from "next/navigation";
import { isAuthenticated } from "@/lib/auth";
import SettingsClient from "./ui";
export default async function SettingsPage(){ if(!(await isAuthenticated())) redirect("/login"); return <SettingsClient/>; }
