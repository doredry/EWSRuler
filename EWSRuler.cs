using FluentArgs;
using Microsoft.Exchange.WebServices.Data;
using System;

namespace EWSInboxRuleGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            FluentArgsBuilder.New()
                .DefaultConfigsWithAppDescription("Simple CLI to generate Exchange inbox rule.")
                .Parameter("-e", "--email")
                    .WithDescription("email address")
                    .WithExamples("test@test.com")
                    .WithValidation(n => !string.IsNullOrWhiteSpace(n), "An email must not only contain whitespace.")
                    .IsRequired()
                .Parameter("-p", "--password")
                    .WithDescription("cleartext password")
                    .WithValidation(n => !string.IsNullOrWhiteSpace(n), "A password must not only contain whitespace.")
                    .WithExamples("123456")
                    .IsRequired()
                 .Parameter("-n", "--name")
                    .WithDescription("Rule name")
                    .WithValidation(n => !string.IsNullOrWhiteSpace(n), "A name must not only contain whitespace.")
                    .IsRequired()
                 .Parameter("-t", "--type")
                    .WithDescription("Forward (forward email to other email) / Delete (Delete incoming emails) / Hide (Move emails to junk folder)")
                    .WithValidation(n => (n == "Forward" || n == "Delete" || n == "Hide"), "type must be one of the following values: Forward \\ Delete \\ Hide")
                    .IsRequired()
                 .ListParameter("-bf", "--bodyfilter")
                    .WithDescription("\"Body contains\" keywords filter")
                    .WithExamples("spam;phish;invoice")
                    .IsOptionalWithEmptyDefault()
                 .ListParameter("-sf", "--subjectfilter")
                    .WithDescription("\"Subject contains\" keywords filter")
                    .WithExamples("spam;phish;invoice")
                    .IsOptionalWithEmptyDefault()
                .Parameter("-r", "--recipient")
                    .WithDescription("Email address to forward the emails to (in case rule type is \"Forward\"")
                    .WithExamples("other@test.com")
                    .IsOptionalWithDefault("null")
                .ListParameter("-f", "--from")
                    .WithDescription("filter by specific senders")
                    .WithExamples("other@test.com;other1@test.com")
                    .IsOptionalWithEmptyDefault()
                .Call(fromsenders => target => subjectfilters => bodyfilters => type => name => password => email =>
                {
                    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
                    service.Credentials = new WebCredentials(email, password);
                    service.AutodiscoverUrl(email);

                    Rule newRule = new Rule();
                    newRule.DisplayName = name;
                    newRule.Priority = 1;
                    newRule.IsEnabled = true;

                    Console.WriteLine($"[V] Authenticated successfuly into {email}\n");

                    if (type == "Forward")
                    {
                        if (target == "null")
                        {
                            Console.WriteLine("[!] Please provide email box to forward the alerts to (by using -r flag)\n");
                            return;
                        }
                        newRule.Actions.ForwardToRecipients.Add(target);
                    }

                    if (type == "Delete")
                    {
                        newRule.Actions.Delete = true;
                        newRule.Actions.MarkAsRead = true;
                    }

                    if (type == "Hide")
                    {
                        newRule.Actions.MarkAsRead = true;
                        newRule.Actions.MoveToFolder = WellKnownFolderName.JunkEmail;
                    }

                    if (bodyfilters.Count > 0)
                    {
                        foreach (var filter in bodyfilters)
                        {
                            newRule.Conditions.ContainsBodyStrings.Add(filter);
                        }
                    }

                    if (subjectfilters.Count > 0)
                    {
                        foreach (var filter in subjectfilters)
                        {
                            newRule.Conditions.ContainsSubjectStrings.Add(filter);
                        }
                    }

                    if (fromsenders.Count > 0)
                    {
                        foreach (var filter in fromsenders)
                        {
                            newRule.Conditions.FromAddresses.Add(filter);
                        }
                    }


                    CreateRuleOperation createOperation = new CreateRuleOperation(newRule);
                    service.UpdateInboxRules(new RuleOperation[] { createOperation }, true);

                    Console.WriteLine($"[V] New inbox rule ({name}) was created!\n");
                    return;

                })
                .Parse(args);
        }
    }
}
