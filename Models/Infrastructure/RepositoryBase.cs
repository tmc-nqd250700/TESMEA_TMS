namespace TESMEA_TMS.Models.Infrastructure
{
    public abstract class RepositoryBase
    {
        protected readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(60); // Timeout mặc định 30s

        /// <summary>
        /// Thực thi một thao tác async với timeout và cancellation token
        /// </summary>
        protected async Task<T> ExecuteAsync<T>(Func<CancellationToken, Task<T>> operation,
                                                CancellationToken cancellationToken = default)
        {
            using var timeoutCts = new CancellationTokenSource(DefaultTimeout);
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, cancellationToken);

            try
            {
                return await operation(linkedCts.Token);
            }
            catch (OperationCanceledException ex)
            {
                if (timeoutCts.IsCancellationRequested)
                {
                    throw new TimeoutException($"Thao tác đã vượt quá thời gian cho phép ({DefaultTimeout.TotalSeconds} giây).", ex);
                }

                throw new OperationCanceledException("Thao tác đã bị hủy bởi người dùng.", ex);
            }
        }

        /// <summary>
        /// Thực thi một thao tác async với timeout (không cần cancellation token)
        /// </summary>
        protected async Task<T> ExecuteAsync<T>(Func<Task<T>> operation)
        {
            using var timeoutCts = new CancellationTokenSource(DefaultTimeout);

            try
            {
                return await operation();
            }
            catch (OperationCanceledException ex)
            {
                if (timeoutCts.IsCancellationRequested)
                {
                    throw new TimeoutException($"Thao tác đã vượt quá thời gian cho phép ({DefaultTimeout.TotalSeconds} giây).", ex);
                }

                throw new OperationCanceledException("Thao tác đã bị hủy bởi hệ thống.", ex);
            }
        }

        /// <summary>
        /// Thực thi một thao tác async không trả về kết quả, với timeout và cancellation token
        /// </summary>
        protected async Task ExecuteAsync(Func<CancellationToken, Task> operation,
                                          CancellationToken cancellationToken = default)
        {
            using var timeoutCts = new CancellationTokenSource(DefaultTimeout);
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, cancellationToken);

            try
            {
                await operation(linkedCts.Token);
            }
            catch (OperationCanceledException ex)
            {
                if (timeoutCts.IsCancellationRequested)
                {
                    throw new TimeoutException($"Thao tác đã vượt quá thời gian cho phép ({DefaultTimeout.TotalSeconds} giây).", ex);
                }

                throw new OperationCanceledException("Thao tác đã bị hủy bởi người dùng.", ex);
            }
        }

        /// <summary>
        /// Thực thi một thao tác async không trả về kết quả, chỉ với timeout (không cần cancellation token)
        /// </summary>
        protected async Task ExecuteAsync(Func<Task> operation)
        {
            using var timeoutCts = new CancellationTokenSource(DefaultTimeout);

            try
            {
                await operation();
            }
            catch (OperationCanceledException ex)
            {
                if (timeoutCts.IsCancellationRequested)
                {
                    throw new TimeoutException($"Thao tác đã vượt quá thời gian cho phép ({DefaultTimeout.TotalSeconds} giây).", ex);
                }

                throw new OperationCanceledException("Thao tác đã bị hủy bởi hệ thống.", ex);
            }
        }
    }
}
