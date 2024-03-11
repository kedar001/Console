using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MyAPI.Controllers
{
    public class MathsController : Controller
    {
        // GET: MathsController
        public ActionResult Index()
        {
            return View();
        }

        // GET: MathsController/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MathsController/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: MathsController/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: MathsController/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MathsController/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: MathsController/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: MathsController/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }
    }
}
