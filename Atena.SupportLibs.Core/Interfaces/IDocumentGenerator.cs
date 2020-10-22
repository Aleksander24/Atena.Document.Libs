using System;
using System.Collections.Generic;
using System.Text;
using Atena.SupportLibs.Core.Enum;

namespace Atena.SupportLibs.Core.Interfaces
{
    public interface IDocumentGenerator
    {
        string Version { get; }
        string Label { get; }
        DocumentTypeEnum DocumentTypeEnum { get; }
        byte[] Generate();
    }
}
