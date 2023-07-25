Reload[] := Module[
    {
        paclet,
        pacletpath
    },
    Enclose[
        paclet = "DataLakeFramework";
        pacletpath = "C:\\Users\\robert.li\\Desktop\\GIT\\DataLakeFramework";
        PacletUninstall[PacletFind[paclet]];
        PacletDirectoryUnload[pacletpath];
        Confirm[PacletDirectoryLoad[pacletpath]];
        <<DataLakeFramework`;
        DataLakeFramework`PackageScope`LoadPackageScope[];
    ]
]