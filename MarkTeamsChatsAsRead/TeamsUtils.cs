using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Task = System.Threading.Tasks.Task;

namespace MarkTeamsChatsAsRead
{
    internal class TeamsUtils
    {

        public static async Task<List<Chat>> GetAllUserChats(GraphServiceClient graphClient)
        {
            var allChats = new List<Chat>();

            IUserChatsCollectionPage? chats = await graphClient.Me.Chats
                .Request()
                .Expand(x => new { x.Members, x.LastMessagePreview })
                .GetAsync();
            do
            {
                Console.WriteLine("Getting next 20 chats ...");

                foreach (var chat in chats)
                {
                    allChats.Add(chat);
                }
                //);

            } while (chats != null && chats.NextPageRequest != null && (chats = await chats?.NextPageRequest?.GetAsync())?.Count > 0);

            return allChats;
        }

        public static async Task<List<Chat>> MarkAllUnreadChatsAsRead(GraphServiceClient graphClient, List<Chat> allChats,
            string userId, string tenantId, bool readAll = false, int parellelism = 1)
        {
            var identity = new TeamworkUserIdentity()
            {
                Id = userId
            };

            ConcurrentQueue<Chat> chatsRead = new ConcurrentQueue<Chat>();

            await Parallel.ForEachAsync(allChats, new ParallelOptions { MaxDegreeOfParallelism = parellelism }, async (chat, token) =>
            {
                Chat? chatVal;
                if ((chatVal = await MarkChatAsRead(graphClient, chat, identity, tenantId, readAll)) != null)
                {
                    chatsRead.Enqueue(chatVal);
                }

            });

            return chatsRead.ToList();
        }

        static async Task<Chat?> MarkChatAsRead(GraphServiceClient graphClient, Chat chat, TeamworkUserIdentity identity, string tenantId, bool readAll)
        {
            var title = chat.Topic;
            if (string.IsNullOrWhiteSpace(title))
            {
                title = string.Join(",", chat.Members.Select(x => x.DisplayName).ToArray());
            }
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"'{title}' ({chat.Id})");

            if ((chat.Viewpoint.LastMessageReadDateTime < chat.LastUpdatedDateTime.GetValueOrDefault().AddSeconds(0))
                || readAll
                )
            {
                var origColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Chat read: MessageReadDateTime/UpdatedDateTime: {chat.Viewpoint.LastMessageReadDateTime}<{chat.LastUpdatedDateTime.GetValueOrDefault().AddSeconds(0)}");

                Console.ForegroundColor = origColor;

                var maxRetries = 3;
                var tries = 0;
                do
                {
                    try
                    {
                        await graphClient.Chats[chat.Id].MarkChatReadForUser(user: identity, tenantId: tenantId).Request().PostAsync();
                        tries = maxRetries;
                        return chat;

                    }
                    catch (ServiceException ex)
                    {
                        tries++;

                        var sleepMs = 5000;

                        if (ex?.ResponseHeaders?.RetryAfter != null)
                        {
                            sleepMs = (int)ex.ResponseHeaders.RetryAfter?.Delta.GetValueOrDefault().TotalMilliseconds;
                        }
                        Console.WriteLine($"Error {ex.Message} Try again in {sleepMs / 1000} seconds.  Will retry {maxRetries - tries} more times");


                        await Task.Delay(sleepMs);

                    }
                } while (tries < maxRetries);

            }
            else
            {
                var origColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Chat not read: MessageReadDateTime/UpdatedDateTime: {chat.Viewpoint.LastMessageReadDateTime}>={chat.LastUpdatedDateTime.GetValueOrDefault().AddSeconds(0)}");

                Console.ForegroundColor = origColor;
            }

            return null;
        }
    }
}
