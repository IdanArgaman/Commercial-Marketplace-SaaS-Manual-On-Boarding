﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace CommandCenter.Controllers
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using CommandCenter.Marketplace;
    using CommandCenter.Models;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Marketplace.SaaS;

    /// <summary>
    /// Manage mail links.
    /// </summary>
    //[Authorize("CommandCenterAdmin")]
    public class MailLinkController : Controller
    {
        private readonly IMarketplaceProcessor marketplaceProcessor;
        private readonly IMarketplaceSaaSClient marketplaceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="MailLinkController"/> class.
        /// </summary>
        /// <param name="marketplaceProcessor">marketplaceProcessor.</param>
        /// <param name="marketplaceClient">Marketplace API client.</param>
        public MailLinkController(IMarketplaceProcessor marketplaceProcessor, IMarketplaceSaaSClient marketplaceClient)
        {
            this.marketplaceProcessor = marketplaceProcessor;
            this.marketplaceClient = marketplaceClient;
        }

        /// <summary>
        /// Activate link.
        /// </summary>
        /// <param name="notificationModel">Details on the URL.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Action result.</returns>
        [HttpGet]
        public async Task<IActionResult> Activate(
            NotificationModel notificationModel,
            CancellationToken cancellationToken)
        {
            if (notificationModel == null)
            {
                throw new ArgumentNullException(nameof(notificationModel));
            }

            var subscriptionDetails = (await this.marketplaceClient.Fulfillment.GetSubscriptionAsync(notificationModel.SubscriptionId, null, null, cancellationToken)).Value;

            if (subscriptionDetails.SaasSubscriptionStatus != Microsoft.Marketplace.SaaS.Models.SubscriptionStatusEnum.PendingFulfillmentStart)
            {
                return this.View(
                    new ActivateActionViewModel
                    {
                        SubscriptionId = notificationModel.SubscriptionId,
                        PlanId = notificationModel.PlanId,
                        Message = $"Could not activate,subscription is in {Enum.GetName(typeof(Microsoft.Marketplace.SaaS.Models.SubscriptionStatusEnum), subscriptionDetails.SaasSubscriptionStatus)} state",
                    });
            }

            await this.marketplaceProcessor.ActivateSubscriptionAsync(notificationModel.SubscriptionId, notificationModel.PlanId, cancellationToken);

            return this.View(
                new ActivateActionViewModel
                {
                    SubscriptionId = notificationModel.SubscriptionId,
                    PlanId = notificationModel.PlanId,
                    Message = "Subscription activated successfully",
                });
        }

        /// <summary>
        /// Quantity change link.
        /// </summary>
        /// <param name="notificationModel">Details on the URL.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Action result.</returns>
        [HttpGet]
        public async Task<IActionResult> QuantityChange(
            NotificationModel notificationModel,
            CancellationToken cancellationToken)
        {
            if (notificationModel == null)
            {
                throw new ArgumentNullException(nameof(notificationModel));
            }

            await this.OperationAckAsync(notificationModel, cancellationToken);

            return this.View("OperationUpdate", notificationModel);
        }

        /// <summary>
        /// Reinstate link.
        /// </summary>
        /// <param name="notificationModel">Details on the URL.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Action result.</returns>
        [HttpGet]
        public async Task<IActionResult> Reinstate(
            NotificationModel notificationModel,
            CancellationToken cancellationToken)
        {
            if (notificationModel == null)
            {
                throw new ArgumentNullException(nameof(notificationModel));
            }

            await this.OperationAckAsync(notificationModel, cancellationToken);

            return this.View("OperationUpdate", notificationModel);
        }

        /// <summary>
        /// Suspend link.
        /// </summary>
        /// <param name="notificationModel">Details on the URL.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Action result.</returns>
        [HttpGet]
        public async Task<IActionResult> SuspendSubscription(
            NotificationModel notificationModel,
            CancellationToken cancellationToken)
        {
            if (notificationModel == null)
            {
                throw new ArgumentNullException(nameof(notificationModel));
            }

            await this.OperationAckAsync(notificationModel, cancellationToken);

            return this.View("OperationUpdate", notificationModel);
        }

        /// <summary>
        /// Unsubscribe link.
        /// </summary>
        /// <param name="notificationModel">Details on the URL.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Action result.</returns>
        [HttpGet]
        public async Task<IActionResult> Unsubscribe(
            NotificationModel notificationModel,
            CancellationToken cancellationToken)
        {
            if (notificationModel == null)
            {
                throw new ArgumentNullException(nameof(notificationModel));
            }

            await this.OperationAckAsync(notificationModel, cancellationToken);

            return this.View("OperationUpdate", notificationModel);
        }

        /// <summary>
        /// Update link.
        /// </summary>
        /// <param name="notificationModel">Details on the URL.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Action result.</returns>
        [HttpGet]
        public async Task<IActionResult> Update(NotificationModel notificationModel, CancellationToken cancellationToken)
        {
            if (notificationModel == null)
            {
                throw new ArgumentNullException(nameof(notificationModel));
            }

            var result = await this.marketplaceClient.Fulfillment.UpdateSubscriptionAsync(
                notificationModel.SubscriptionId,
                new Microsoft.Marketplace.SaaS.Models.SubscriberPlan { PlanId = notificationModel.PlanId },
                null,
                null,
                cancellationToken);

            return this.View(
                new ActivateActionViewModel
                {
                    SubscriptionId = notificationModel.SubscriptionId,
                    PlanId = notificationModel.PlanId,
                });
        }

        private async Task OperationAckAsync(
            NotificationModel payload,
            CancellationToken cancellationToken)
        {
            await this.marketplaceProcessor.OperationAckAsync(payload.SubscriptionId, payload.OperationId, payload.PlanId, payload.Quantity, cancellationToken);
        }
    }
}
