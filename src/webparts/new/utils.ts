export const nameof = <T extends {}>(name: keyof T) => name;
export const sleep = (ms:any) => new Promise((resolve) => setTimeout(resolve, ms));
