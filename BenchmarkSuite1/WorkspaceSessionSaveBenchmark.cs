using System;
using System.IO;
using System.Reflection;
using System.Threading;
using BenchmarkDotNet.Attributes;
using Microsoft.VSDiagnostics;

namespace LumiCanvas.Benchmarks;
[CPUUsageDiagnoser]
public class WorkspaceSessionSaveBenchmark
{
    private WorkspaceSession _session = null!;
    private TaskBoard _task = null!;
    private MethodInfo _createSnapshotMethod = null!;
    private MethodInfo _writeArchiveMethod = null!;
    [GlobalSetup]
    public void GlobalSetup()
    {
        _session = new WorkspaceSession();
        _task = _session.AddTask("BenchmarkTask");
        var assetsRoot = _session.GetTaskAssetsDirectory(_task.Id);
        for (var i = 0; i < 20; i++)
        {
            var filePath = Path.Combine(assetsRoot, $"asset_{i}.bin");
            if (!File.Exists(filePath))
            {
                File.WriteAllBytes(filePath, new byte[64 * 1024]);
            }

            _task.Items.Add(new BoardItemModel { Kind = BoardItemKind.File, X = i * 10, Y = i * 10, Width = 240, Height = 80, SourcePath = filePath });
        }

        var sessionType = typeof(WorkspaceSession);
        _createSnapshotMethod = sessionType.GetMethod("CreateTaskSaveSnapshot", BindingFlags.Instance | BindingFlags.NonPublic) ?? throw new InvalidOperationException("CreateTaskSaveSnapshot not found.");
        _writeArchiveMethod = sessionType.GetMethod("WriteTaskArchive", BindingFlags.Instance | BindingFlags.NonPublic) ?? throw new InvalidOperationException("WriteTaskArchive not found.");
    }

    [Benchmark]
    public void WriteTaskArchive()
    {
        var snapshot = _createSnapshotMethod.Invoke(_session, new object[] { _task }) ?? throw new InvalidOperationException("Failed to create snapshot.");
        _writeArchiveMethod.Invoke(_session, new object[] { snapshot, CancellationToken.None });
    }
}