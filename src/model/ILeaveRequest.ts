import { LeaveType } from "./LeaveType";

export interface ILeaveRequest {
    username: string;
    type: LeaveType;
    startDate: Date;
    finishDate: Date;
    calculatedHours: number;
    approved: boolean;
}