﻿namespace OfficeDevPnP.Core.Framework.Authentication.Events
{
    using Microsoft.AspNet.Http;

    public class AuthenticationSucceededContext : BaseSharePointAuthenticationContext
    {
        public AuthenticationSucceededContext(HttpContext context, SharePointAuthenticationOptions options)
               : base(context, options)
        {
        }

        public SharePointContext SharePointContext { get; set; }
    }
}
