// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace CommandCenter.Controllers
{
    using System;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading;
    using System.Threading.Tasks;
    using CommandCenter.Authorization;
    using CommandCenter.Marketplace;
    using CommandCenter.Models;
    using Microsoft.AspNetCore.Authentication.OpenIdConnect;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Microsoft.Identity.Web;
    using Microsoft.Marketplace.SaaS;
    using Microsoft.Marketplace.SaaS.Models;

    /// <summary>
    /// Landing page.
    /// </summary>
    //[Authorize(AuthenticationSchemes = OpenIdConnectDefaults.AuthenticationScheme)] // Specify the auth scheme to be used for logging on users. This is for supporting WebAPI auth
     [AllowAnonymous]
    public class LandingPageController : Controller
    {
        private const string SampleToken = "sampletoken";
        private readonly ILogger<LandingPageController> logger;
        private readonly IMarketplaceProcessor marketplaceProcessor;
        private readonly IMarketplaceNotificationHandler notificationHandler;
        private readonly IMarketplaceSaaSClient marketplaceClient;
        private readonly GraphServiceClient graphServiceClient;
        private readonly CommandCenterOptions options;

        /// <summary>
        /// Initializes a new instance of the <see cref="LandingPageController"/> class.
        /// </summary>
        /// <param name="commandCenterOptions">Options.</param>
        /// <param name="marketplaceProcessor">Marketplace processor.</param>
        /// <param name="notificationHandler">Notification handler.</param>
        /// <param name="marketplaceClient">Marketplace client.</param>
        /// <param name="graphServiceClient">Client for Graph API.</param>
        /// <param name="logger">Logger.</param>
        public LandingPageController(
            //IOptionsMonitor<CommandCenterOptions> commandCenterOptions,
            // IMarketplaceProcessor marketplaceProcessor,
            // IMarketplaceNotificationHandler notificationHandler,
            // IMarketplaceSaaSClient marketplaceClient,
            //GraphServiceClient graphServiceClient,
            //ILogger<LandingPageController> logger
        )
        {
            // if (commandCenterOptions == null)
            // {
            //     throw new ArgumentNullException(nameof(commandCenterOptions));
            // }

            // this.marketplaceProcessor = marketplaceProcessor;
            // this.notificationHandler = notificationHandler;
            // this.marketplaceClient = marketplaceClient;
            // this.graphServiceClient = graphServiceClient;
            // this.logger = logger;
            // this.options = commandCenterOptions.CurrentValue;
        }

        /// <summary>
        /// Landing page get.
        /// </summary>
        /// <param name="token">Marketplace purchase identification token.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Action result.</returns>
        [AuthorizeForScopes(Scopes = new string[] { "user.read" })]
        public async Task<ActionResult> Index(string token, CancellationToken cancellationToken)
        {
            /*
            if (string.IsNullOrEmpty(token))
            {
                this.ModelState.AddModelError(string.Empty, "Token URL parameter cannot be empty");
                this.ViewBag.Message = "Token URL parameter cannot be empty";
                return this.View();
            }

            ResolvedSubscription resolvedSubscription = null;
            Microsoft.Marketplace.SaaS.Models.Subscription subscriptionDetails = null;
            Azure.Response<Microsoft.Marketplace.SaaS.Models.SubscriptionPlans> availablePlans = null;
            bool anyPendingOperations = false;

            if (token.ToLowerInvariant() != SampleToken)
            {
                // Get the subscription for the offer from the marketplace purchase identification token
                resolvedSubscription = await this.marketplaceProcessor.GetSubscriptionFromPurchaseIdentificationTokenAsync(token, cancellationToken);

                if (resolvedSubscription == default(ResolvedSubscription))
                {
                    this.ViewBag.Message = "Token did not resolve to a subscription";
                    return this.View();
                }

                subscriptionDetails = resolvedSubscription.Subscription;

                // Populate the available plans for this subscription from the API
                availablePlans = await this.marketplaceClient.Fulfillment.ListAvailablePlansAsync(
                    resolvedSubscription.Id.Value,
                    null,
                    null,
                    cancellationToken);

                // See if there are pending operations for this subscription
                var pendingOperations = await this.marketplaceClient.Operations.ListOperationsAsync(
                    resolvedSubscription.Id.Value,
                    null,
                    null,
                    cancellationToken);
                anyPendingOperations = pendingOperations?.Value.Operations?.Any(o => o.Status == OperationStatusEnum.InProgress) ?? false;
            }

            var graphApiUser = await this.graphServiceClient.Me.Request().GetAsync();

            var isSampleToken = string.Equals(token, SampleToken, StringComparison.InvariantCultureIgnoreCase);
            */

            var isSampleToken = true;
            var provisioningModel = new AzureSubscriptionProvisionModel
            {
                // Landing page is the only place to capture the customer's contact details
                // It can be present in multiple places:
                //  - the details received from the Graph API
                //  - beneficiary information on the subscription details
                // it is also possible that the Graph API
                NameFromOpenIdConnect = "Idan", // (this.User.Identity as ClaimsIdentity)?.FindFirst("name")?.Value,
                EmailFromClaims = "Idan", //this.User.Identity.GetUserEmail(),
                EmailFromGraph = "Idan", //graphApiUser.Mail ?? string.Empty,
                NameFromGraph = "Idan", //graphApiUser.DisplayName ?? string.Empty,
                UserPrincipalName = "Idan", //graphApiUser.UserPrincipalName ?? string.Empty,
                PurchaserEmail = "Idan", //graphApiUser.Mail ?? string.Empty,

                // Get the other potential contact information from the marketplace API
                PurchaserUPN = "purchaser@purchaser.com", //isSampleToken ? "purchaser@purchaser.com" : subscriptionDetails?.Purchaser?.EmailId,
                PurchaserTenantId = Guid.Empty, //isSampleToken ? Guid.Empty : subscriptionDetails?.Purchaser?.TenantId ?? Guid.Empty,
                BeneficiaryUPN = "purchaser@purchaser.com", //isSampleToken ? "customer@customer.com" : subscriptionDetails?.Beneficiary?.EmailId,
                BeneficiaryTenantId = Guid.Empty, //isSampleToken ? Guid.Empty : subscriptionDetails?.Beneficiary?.TenantId ?? Guid.Empty,

                // Maybe the end users are a completely different set of contacts, start with one
                BusinessUnitContactEmail = "Idan", //this.User.Identity.GetUserEmail(),

                PlanId = "purchaser@purchaser.com", //isSampleToken ? "purchaser@purchaser.com" : resolvedSubscription.PlanId,
                SubscriptionId =  Guid.Empty, //isSampleToken ? Guid.Empty : resolvedSubscription.Id.Value,
                OfferId = "purchaser@purchaser.com", //isSampleToken ? "sample offer" : resolvedSubscription.OfferId,
                SubscriptionName = "sample subscription", //isSampleToken ? "sample subscription" : resolvedSubscription.SubscriptionName,
                SubscriptionStatus = SubscriptionStatusEnum.PendingFulfillmentStart, //isSampleToken ? SubscriptionStatusEnum.PendingFulfillmentStart : subscriptionDetails?.SaasSubscriptionStatus ?? SubscriptionStatusEnum.NotStarted,

                Region = TargetContosoRegionEnum.NorthAmerica,
                AvailablePlans = new System.Collections.Generic.List<Plan>(), //isSampleToken ? new System.Collections.Generic.List<Plan>() : availablePlans?.Value.Plans.ToList(),
                PendingOperations = false, //isSampleToken ? false : anyPendingOperations,
            };

            return this.View(provisioningModel);
        }

        /// <summary>
        /// Landing page post handler.
        /// </summary>
        /// <param name="provisionModel">View model.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Action result.</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Index(AzureSubscriptionProvisionModel provisionModel, CancellationToken cancellationToken)
        {
            if (provisionModel == null)
            {
                throw new ArgumentNullException(nameof(provisionModel));
            }

            var urlBase = $"{this.Request.Scheme}://{this.Request.Host}";
            this.options.BaseUrl = new Uri(urlBase);
            try
            {
                // A new subscription will have PendingFulfillmentStart as status
                if (provisionModel.SubscriptionStatus == SubscriptionStatusEnum.PendingFulfillmentStart)
                {
                    await this.notificationHandler.ProcessNewSubscriptionAsyc(provisionModel, cancellationToken);
                }
                else
                {
                    await this.notificationHandler.ProcessChangePlanAsync(provisionModel, cancellationToken);
                }

                return this.RedirectToAction(nameof(this.Success));
            }
            catch (Exception ex)
            {
                return this.BadRequest(ex);
            }
        }

        /// <summary>
        /// Success.
        /// </summary>
        /// <returns>Action result.</returns>
        public ActionResult Success()
        {
            return this.View();
        }
    }
}
