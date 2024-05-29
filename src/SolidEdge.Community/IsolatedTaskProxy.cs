using System;
using System.Threading;
using SolidEdgeFramework;

namespace SolidEdgeCommunity;

/// <summary>
///     Abstract base class to be used with IsolatedTask&lt;T&gt;.
/// </summary>
public abstract class IsolatedTaskProxy : MarshalByRefObject
{
    private Application _application;
    private SolidEdgeDocument _document;

    /// <summary>
    ///     Solid Edge Application property.
    /// </summary>
    public Application Application
    {
        get => _application;
        set => _application = UnwrapRuntimeCallableWrapper<Application>(value);
    }

    /// <summary>
    ///     Solid Edge Application property.
    /// </summary>
    public SolidEdgeDocument Document
    {
        get => _document;
        set => _document = UnwrapRuntimeCallableWrapper<SolidEdgeDocument>(value);
    }

    /// <summary>
    ///     Lifetime services as disabled by default.
    /// </summary>
    [Obsolete("This Remoting API is not supported and throws PlatformNotSupportedException.",
        DiagnosticId = "SYSLIB0010", UrlFormat = "https://aka.ms/dotnet-warnings/{0}")]
    public sealed override object InitializeLifetimeService()
    {
        return null;
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <param name="target"></param>
    protected void InvokeSTAThread(Action target)
    {
        if (target == null) throw new ArgumentNullException("target");

        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                target();
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TArg1">The type of arg1.</typeparam>
    /// <param name="target"></param>
    /// <param name="arg1"></param>
    protected void InvokeSTAThread<TArg1>(Action<TArg1> target, TArg1 arg1)
    {
        if (target == null) throw new ArgumentNullException("target");

        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                target(arg1);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TArg1">The type of arg1.</typeparam>
    /// <typeparam name="TArg2">The type of arg2.</typeparam>
    /// <param name="target"></param>
    /// <param name="arg1"></param>
    /// <param name="arg2"></param>
    protected void InvokeSTAThread<TArg1, TArg2>(Action<TArg1, TArg2> target, TArg1 arg1, TArg2 arg2)
    {
        if (target == null) throw new ArgumentNullException("target");

        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                target(arg1, arg2);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TArg1">The type of arg1.</typeparam>
    /// <typeparam name="TArg2">The type of arg2.</typeparam>
    /// <typeparam name="TArg3">The type of arg3.</typeparam>
    /// <param name="target"></param>
    /// <param name="arg1"></param>
    /// <param name="arg2"></param>
    /// <param name="arg3"></param>
    protected void InvokeSTAThread<TArg1, TArg2, TArg3>(Action<TArg1, TArg2, TArg3> target, TArg1 arg1, TArg2 arg2,
        TArg3 arg3)
    {
        if (target == null) throw new ArgumentNullException("target");

        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                target(arg1, arg2, arg3);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TArg1">The type of arg1.</typeparam>
    /// <typeparam name="TArg2">The type of arg2.</typeparam>
    /// <typeparam name="TArg3">The type of arg3.</typeparam>
    /// <typeparam name="TArg4">The type of arg4.</typeparam>
    /// <param name="target"></param>
    /// <param name="arg1"></param>
    /// <param name="arg2"></param>
    /// <param name="arg3"></param>
    /// <param name="arg4"></param>
    protected void InvokeSTAThread<TArg1, TArg2, TArg3, TArg4>(Action<TArg1, TArg2, TArg3, TArg4> target, TArg1 arg1,
        TArg2 arg2, TArg3 arg3, TArg4 arg4)
    {
        if (target == null) throw new ArgumentNullException("target");

        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                target(arg1, arg2, arg3, arg4);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TResult">The type of the return value.</typeparam>
    /// <param name="target"></param>
    /// <returns>An instance of TResult.</returns>
    protected TResult InvokeSTAThread<TResult>(Func<TResult> target)
    {
        if (target == null) throw new ArgumentNullException("target");

        var returnValue = default(TResult);
        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                returnValue = target();
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);

        return returnValue;
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TArg1">The type of arg1.</typeparam>
    /// <typeparam name="TResult">The type of the return value.</typeparam>
    /// <param name="target"></param>
    /// <param name="arg1"></param>
    /// <returns>An instance of TResult.</returns>
    protected TResult InvokeSTAThread<TArg1, TResult>(Func<TArg1, TResult> target, TArg1 arg1)
    {
        if (target == null) throw new ArgumentNullException("target");

        var returnValue = default(TResult);
        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                returnValue = target(arg1);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);

        return returnValue;
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TArg1">The type of arg1.</typeparam>
    /// <typeparam name="TArg2">The type of arg2.</typeparam>
    /// <typeparam name="TResult">The type of the return value.</typeparam>
    /// <param name="target"></param>
    /// <param name="arg1"></param>
    /// <param name="arg2"></param>
    /// <returns>An instance of TResult.</returns>
    protected TResult InvokeSTAThread<TArg1, TArg2, TResult>(Func<TArg1, TArg2, TResult> target, TArg1 arg1, TArg2 arg2)
    {
        if (target == null) throw new ArgumentNullException("target");

        var returnValue = default(TResult);
        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                returnValue = target(arg1, arg2);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);

        return returnValue;
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TArg1">The type of arg1.</typeparam>
    /// <typeparam name="TArg2">The type of arg2.</typeparam>
    /// <typeparam name="TArg3">The type of arg3.</typeparam>
    /// <typeparam name="TResult">The type of the return value.</typeparam>
    /// <param name="target"></param>
    /// <param name="arg1"></param>
    /// <param name="arg2"></param>
    /// <param name="arg3"></param>
    /// <returns>An instance of TResult.</returns>
    protected TResult InvokeSTAThread<TArg1, TArg2, TArg3, TResult>(Func<TArg1, TArg2, TArg3, TResult> target,
        TArg1 arg1, TArg2 arg2, TArg3 arg3)
    {
        if (target == null) throw new ArgumentNullException("target");

        var returnValue = default(TResult);
        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                returnValue = target(arg1, arg2, arg3);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);

        return returnValue;
    }

    /// <summary>
    ///     Invokes a method in a STA thread.
    /// </summary>
    /// <typeparam name="TArg1">The type of arg1.</typeparam>
    /// <typeparam name="TArg2">The type of arg2.</typeparam>
    /// <typeparam name="TArg3">The type of arg3.</typeparam>
    /// <typeparam name="TArg4">The type of arg4.</typeparam>
    /// <typeparam name="TResult">The type of the return value.</typeparam>
    /// <param name="target"></param>
    /// <param name="arg1"></param>
    /// <param name="arg2"></param>
    /// <param name="arg3"></param>
    /// <param name="arg4"></param>
    /// <returns>An instance of TResult.</returns>
    protected TResult InvokeSTAThread<TArg1, TArg2, TArg3, TArg4, TResult>(
        Func<TArg1, TArg2, TArg3, TArg4, TResult> target, TArg1 arg1, TArg2 arg2, TArg3 arg3, TArg4 arg4)
    {
        if (target == null) throw new ArgumentNullException("target");

        var returnValue = default(TResult);
        Exception exception = null;

        // Define thread.
        var thread = new Thread(() =>
        {
            // Thread specific try\catch.
            try
            {
                returnValue = target(arg1, arg2, arg3, arg4);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        // Important! Set thread apartment state to STA.
        thread.SetApartmentState(ApartmentState.STA);

        // Start the thread.
        thread.Start();

        // Wait for the thead to finish.
        thread.Join();

        if (exception != null)
            throw new Exception("An unhandled exception has occurred. See inner exception for details.", exception);

        return returnValue;
    }

    /// <summary>
    ///     Unwraps a runtime callable wrapper (RCW) that is passed across AppDomains.
    /// </summary>
    /// <typeparam name="TInterface"></typeparam>
    /// <param name="rcw"></param>
    /// <returns></returns>
    protected TInterface UnwrapRuntimeCallableWrapper<TInterface>(object rcw) where TInterface : class
    {
        /*if (RemotingServices.IsTransparentProxy(rcw))
        {
            if (System.Runtime.InteropServices.Marshal.IsComObject(rcw))
            {
                IntPtr punk = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(rcw);

                try
                {
                    return (TInterface)System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(punk);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.Release(punk);
                }
            }
            else
            {
                throw new InvalidComObjectException();
            }
        }*/

        return rcw as TInterface;
    }
}