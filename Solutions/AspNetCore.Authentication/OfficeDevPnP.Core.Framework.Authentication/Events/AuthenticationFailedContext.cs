﻿namespace OfficeDevPnP.Core.Framework.Authentication.Events
{
    using Microsoft.AspNet.Http;
    public class AuthenticationFailedContext : BaseSharePointAuthenticationContext
    {
        public AuthenticationFailedContext(HttpContext context, SharePointAuthenticationOptions options)
               : base(context, options)
        {
        }
    }
}