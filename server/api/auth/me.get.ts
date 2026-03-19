import { defineEventHandler } from "h3";
import { getOptionalUser } from "../../utils/auth";

export default defineEventHandler(async (event) => {
  const user = await getOptionalUser(event);
  return {
    authenticated: Boolean(user),
    user
  };
});
